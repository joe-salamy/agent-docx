import { randomUUID } from "node:crypto";
import {
  link,
  lstat,
  mkdir,
  readFile,
  readdir,
  realpath,
  rename,
  rm,
} from "node:fs/promises";
import {
  basename,
  dirname,
  isAbsolute,
  relative,
  resolve,
  sep,
} from "node:path";
import { AgentDocxError } from "../../types.js";
import { objectRecord } from "../../json-contract.js";
import { pathExists, writeAtomicFile, writeExclusiveFile } from "../fs-util.js";
import { isDocumentId, type RevisionId } from "../../legal/model.js";
import type {
  AgentDocxDocumentConfig,
  AgentDocxManifest,
  ProjectDocumentInput,
} from "../contracts.js";
import {
  BINDING_FILE,
  EXPORT_INTENT,
  INIT_INTENT,
  LOCK_NAME,
  STORE_DIR,
  type ExportIntent,
  type OpenedStore,
} from "../store.js";
import {
  assertNoSymlinkComponents,
  assertDirectory,
  assertRegularFile,
  assertRelativeManifestPath,
  assertWithin,
  canonicalJson,
  emptyObjectId,
  objectId,
  readJsonFile,
  readProjectFile,
  relativePath,
  strictJson,
} from "./objects.js";
import { validateManifest, validateManifestPaths } from "./validate.js";

export { validateManifest, validateManifestPaths } from "./validate.js";
export {
  assertDirectory,
  assertRegularFile,
  readProjectFile,
  strictJson,
} from "./objects.js";

type ProjectBinding = {
  schemaVersion: 1;
  projectId: string;
  manifestBasename: string;
};

type InitializationIntent = {
  schemaVersion: 1;
  state: "preparing";
  projectId: string;
  manifestPath: string;
  storePath: string;
};

const bindingPath = (storePath: string): string =>
  resolve(storePath, BINDING_FILE);

export const validateBinding = async (opened: OpenedStore): Promise<void> => {
  const binding = await readJsonFile<ProjectBinding>(
    bindingPath(opened.storePath),
  );
  if (
    binding.schemaVersion !== 1 ||
    binding.projectId !== opened.manifest.projectId ||
    binding.manifestBasename !== basename(opened.manifestPath)
  )
    throw new AgentDocxError(
      "PROJECT_INVALID",
      "Manifest does not own this project store",
    );
};

export const recoverInitialization = async (
  projectDirectory: string,
  manifestPath: string,
): Promise<void> => {
  const intentPath = resolve(projectDirectory, INIT_INTENT);
  let intent: InitializationIntent;
  try {
    intent = await readJsonFile<InitializationIntent>(intentPath);
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === "ENOENT") return;
    throw error;
  }
  if (
    intent.schemaVersion !== 1 ||
    intent.state !== "preparing" ||
    intent.manifestPath !== manifestPath ||
    intent.storePath !== resolve(projectDirectory, STORE_DIR)
  )
    throw new AgentDocxError(
      "PROJECT_INVALID",
      "Malformed initialization intent",
    );
  const storePath = intent.storePath;

  let committed = false;
  try {
    const manifest = validateManifest(
      await readJsonFile<unknown>(manifestPath),
    );
    const binding = await readJsonFile<ProjectBinding>(bindingPath(storePath));
    committed =
      manifest.projectId === intent.projectId &&
      binding.projectId === intent.projectId &&
      binding.manifestBasename === basename(manifestPath);
  } catch (error) {
    const code =
      error instanceof AgentDocxError
        ? error.code
        : (error as NodeJS.ErrnoException).code;
    const partialInitialization =
      code === "INPUT_NOT_FOUND" ||
      code === "PROJECT_INVALID" ||
      code === "ENOENT" ||
      code === "ENOTDIR";
    if (!partialInitialization) throw error;
  }
  if (committed) {
    await rm(intentPath, { force: true });
    return;
  }
  const binding = await readJsonFile<ProjectBinding>(
    bindingPath(storePath),
  ).catch(() => null);
  if (binding?.projectId === intent.projectId) {
    await rm(storePath, { recursive: true, force: true });
    await rm(manifestPath, { force: true });
  }
  await rm(intentPath, { force: true });
};
const UUID_SUFFIX =
  "[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}";
const staleStageFilePattern = new RegExp(
  `\\.${UUID_SUFFIX}\\.(?:stage|backup)$`,
);
const exportStageDirectoryPattern = new RegExp(
  `\\.agent-docx-${UUID_SUFFIX}\\.stage$`,
);

export const sweepOwnedStages = async (
  projectDirectory: string,
  manifestPath: string,
): Promise<void> => {
  const visit = async (directory: string): Promise<void> => {
    let entries;
    try {
      entries = await readdir(directory, { withFileTypes: true });
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code === "ENOENT") return;
      throw error;
    }
    for (const entry of entries.sort((left, right) =>
      left.name < right.name ? -1 : left.name > right.name ? 1 : 0,
    )) {
      const path = resolve(directory, entry.name);
      if (entry.isSymbolicLink()) continue;
      if (entry.isDirectory()) {
        if (exportStageDirectoryPattern.test(entry.name)) {
          try {
            const marker = strictJson<Record<string, unknown>>(
              await readFile(stageMarker(path)),
              stageMarker(path),
            );
            if (
              marker.schemaVersion === 1 &&
              typeof marker.owner === "string" &&
              marker.owner.length > 0 &&
              typeof marker.projectId === "string" &&
              marker.projectId.length > 0 &&
              marker.manifestPath === manifestPath
            )
              await rm(path, { recursive: true, force: true });
          } catch (error) {
            if ((error as NodeJS.ErrnoException).code !== "ENOENT") throw error;
          }
          continue;
        }
        if (
          entry.name === ".git" ||
          entry.name === "node_modules" ||
          entry.name === "dist"
        )
          continue;
        await visit(path);
        continue;
      }
      if (entry.isFile() && staleStageFilePattern.test(entry.name))
        await rm(path, { force: true });
    }
  };
  await visit(projectDirectory);
};

const exportIntentPath = (projectDirectory: string): string =>
  resolve(projectDirectory, EXPORT_INTENT);

const intentPathValue = (
  projectDirectory: string,
  value: unknown,
  label: string,
): string => {
  if (typeof value !== "string" || !isAbsolute(value))
    throw new AgentDocxError(
      "PROJECT_INVALID",
      `Export intent ${label} is invalid`,
    );
  const relativeValue = relative(projectDirectory, value);
  if (
    relativeValue === ".." ||
    relativeValue.startsWith(`..${sep}`) ||
    isAbsolute(relativeValue)
  )
    throw new AgentDocxError(
      "PROJECT_INVALID",
      `Export intent ${label} escapes the project directory`,
    );
  return value;
};

const exportIntentFrom = (
  projectDirectory: string,
  value: unknown,
): ExportIntent => {
  const intent = objectRecord(value, "Export intent", {
    code: "PROJECT_INVALID",
  });
  const state = intent.state;
  if (
    intent.schemaVersion !== 1 ||
    (state !== "preparing" && state !== "prepared") ||
    typeof intent.projectId !== "string" ||
    typeof intent.owner !== "string" ||
    intent.owner.length === 0
  )
    throw new AgentDocxError("PROJECT_INVALID", "Malformed export intent");
  const manifestPath = intentPathValue(
    projectDirectory,
    intent.manifestPath,
    "manifestPath",
  );
  if (manifestPath !== resolve(projectDirectory, basename(manifestPath)))
    throw new AgentDocxError(
      "PROJECT_INVALID",
      "Export intent manifest path is invalid",
    );
  const outputPath = intentPathValue(
    projectDirectory,
    intent.outputPath,
    "outputPath",
  );
  const attachmentPath =
    intent.attachmentPath === null
      ? null
      : intentPathValue(
          projectDirectory,
          intent.attachmentPath,
          "attachmentPath",
        );
  const stagePath = intentPathValue(
    projectDirectory,
    intent.stagePath,
    "stagePath",
  );
  const docxStagePath = intentPathValue(
    projectDirectory,
    intent.docxStagePath,
    "docxStagePath",
  );
  const pdfStagePath =
    intent.pdfStagePath === undefined || intent.pdfStagePath === null
      ? null
      : intentPathValue(projectDirectory, intent.pdfStagePath, "pdfStagePath");
  const attachmentStagePath =
    intent.attachmentStagePath === null
      ? null
      : intentPathValue(
          projectDirectory,
          intent.attachmentStagePath,
          "attachmentStagePath",
        );
  const artifactStorePath = intentPathValue(
    projectDirectory,
    intent.artifactStorePath,
    "artifactStorePath",
  );
  const attachmentStorePath =
    intent.attachmentStorePath === null
      ? null
      : intentPathValue(
          projectDirectory,
          intent.attachmentStorePath,
          "attachmentStorePath",
        );
  const pdfStorePath =
    intent.pdfStorePath === undefined || intent.pdfStorePath === null
      ? null
      : intentPathValue(projectDirectory, intent.pdfStorePath, "pdfStorePath");
  const digest = (value: unknown, label: string): RevisionId => {
    if (typeof value !== "string" || !/^sha256:[0-9a-f]{64}$/.test(value))
      throw new AgentDocxError(
        "PROJECT_INVALID",
        `Export intent ${label} is invalid`,
      );
    return value as RevisionId;
  };
  const artifactProvenanceSha256 = digest(
    intent.artifactProvenanceSha256,
    "artifactProvenanceSha256",
  );
  const pdfSha256 =
    intent.pdfSha256 === undefined || intent.pdfSha256 === null
      ? emptyObjectId
      : digest(intent.pdfSha256, "pdfSha256");
  const artifactsRoot = resolve(projectDirectory, STORE_DIR, "artifacts");
  const artifactRelative = relative(artifactsRoot, dirname(artifactStorePath));
  if (
    artifactRelative.split(sep).length !== 2 ||
    artifactRelative.split(sep).some((part) => !/^[0-9a-f]{64}$/.test(part)) ||
    basename(artifactStorePath) !== "document.docx"
  )
    throw new AgentDocxError(
      "PROJECT_INVALID",
      "Export intent artifactStorePath is invalid",
    );
  if (
    pdfStorePath !== null &&
    (dirname(pdfStorePath) !== dirname(artifactStorePath) ||
      basename(pdfStorePath) !== "document.pdf")
  )
    throw new AgentDocxError(
      "PROJECT_INVALID",
      "Export intent pdfStorePath is invalid",
    );
  if (
    state === "prepared" &&
    (pdfSha256 !== emptyObjectId) !== (pdfStorePath !== null)
  )
    throw new AgentDocxError(
      "PROJECT_INVALID",
      "Prepared PDF export is incomplete",
    );
  if (
    attachmentStorePath !== null &&
    !attachmentStorePath.startsWith(`${dirname(artifactStorePath)}/`) &&
    !attachmentStorePath.startsWith(`${dirname(artifactStorePath)}${sep}`)
  )
    throw new AgentDocxError(
      "PROJECT_INVALID",
      "Export intent attachmentStorePath is invalid",
    );
  if (
    outputPath === resolve(projectDirectory) ||
    !stagePath.startsWith(`${outputPath}.agent-docx-${intent.owner}.stage`) ||
    stagePath !== `${outputPath}.agent-docx-${intent.owner}.stage` ||
    docxStagePath !== resolve(stagePath, "document.docx") ||
    (attachmentStagePath !== null &&
      attachmentStagePath !== resolve(stagePath, "attachments")) ||
    (attachmentPath !== null && attachmentStagePath === null) ||
    (pdfStagePath !== null &&
      pdfStagePath !== resolve(stagePath, "document.pdf")) ||
    (pdfSha256 !== emptyObjectId &&
      (pdfStagePath === null || pdfStorePath === null)) ||
    (pdfSha256 === emptyObjectId && pdfStorePath !== null) ||
    (pdfStagePath !== null &&
      pdfSha256 === emptyObjectId &&
      state === "prepared")
  )
    throw new AgentDocxError(
      "PROJECT_INVALID",
      "Malformed export intent paths",
    );
  return {
    schemaVersion: 1,
    state,
    projectId: intent.projectId,
    manifestPath,
    owner: intent.owner,
    outputPath,
    attachmentPath,
    stagePath,
    docxStagePath,
    pdfStagePath,
    attachmentStagePath,
    docxSha256: digest(intent.docxSha256, "docxSha256"),
    attachmentManifestSha256:
      intent.attachmentManifestSha256 === null
        ? null
        : digest(intent.attachmentManifestSha256, "attachmentManifestSha256"),
    artifactProvenanceSha256,
    artifactStorePath,
    attachmentStorePath,
    pdfStorePath,
    pdfSha256,
  };
};

const stageMarker = (stagePath: string): string =>
  resolve(stagePath, "owner.json");

const ownedStage = async (
  stagePath: string,
  owner: string,
  projectId: string,
  manifestPath: string,
): Promise<boolean> => {
  let entry;
  try {
    entry = await lstat(stagePath);
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === "ENOENT") return false;
    throw error;
  }
  if (!entry.isDirectory() || entry.isSymbolicLink())
    throw new AgentDocxError(
      "PROJECT_INVALID",
      `Export stage is not a directory: ${stagePath}`,
    );
  try {
    const marker = strictJson<Record<string, unknown>>(
      await readFile(stageMarker(stagePath)),
      stageMarker(stagePath),
    );
    return (
      marker.schemaVersion === 1 &&
      marker.owner === owner &&
      marker.projectId === projectId &&
      marker.manifestPath === manifestPath
    );
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === "ENOENT") return false;
    throw error;
  }
};

const removeOwnedStage = async (intent: ExportIntent): Promise<void> => {
  if (
    await ownedStage(
      intent.stagePath,
      intent.owner,
      intent.projectId,
      intent.manifestPath,
    )
  )
    await rm(intent.stagePath, { recursive: true, force: true });
};

const hashRegularFile = async (
  path: string,
  label: string,
): Promise<RevisionId> => {
  await assertRegularFile(path, label);
  return objectId(await readFile(path));
};

const attachmentPayloadPath = (value: unknown): string => {
  if (
    typeof value !== "string" ||
    !value.startsWith("files/") ||
    value.includes("\\") ||
    value
      .split("/")
      .some((part, index) =>
        index === 0
          ? part !== "files"
          : part.length === 0 || part === "." || part === "..",
      )
  )
    throw new AgentDocxError(
      "PROJECT_INVALID",
      "Attachment manifest payload path is invalid",
    );
  return value;
};

const verifyAttachmentDirectory = async (
  directory: string,
  manifestSha256: RevisionId,
): Promise<void> => {
  await assertDirectory(directory, "Published attachment bundle");
  const manifestPath = resolve(directory, "manifest.json");
  await assertRegularFile(manifestPath, "Attachment manifest");
  const manifestBytes = await readFile(manifestPath);
  if (objectId(manifestBytes) !== manifestSha256)
    throw new AgentDocxError(
      "PROJECT_INVALID",
      "Published attachment manifest hash mismatch",
    );
  const manifest = strictJson<{
    schemaVersion: number;
    entries: readonly {
      name: string;
      mediaType: string;
      byteLength: number;
      sha256: string;
      payloadPath: string;
    }[];
  }>(manifestBytes, manifestPath);
  if (
    manifest.schemaVersion !== 1 ||
    !Array.isArray(manifest.entries) ||
    manifest.entries.some(
      (entry) =>
        !entry ||
        typeof entry.name !== "string" ||
        typeof entry.mediaType !== "string" ||
        !Number.isSafeInteger(entry.byteLength) ||
        entry.byteLength < 0 ||
        typeof entry.sha256 !== "string" ||
        !/^sha256:[0-9a-f]{64}$/.test(entry.sha256),
    )
  )
    throw new AgentDocxError(
      "PROJECT_INVALID",
      "Published attachment manifest is invalid",
    );
  const names = new Set<string>();
  for (const entry of manifest.entries) {
    if (names.has(entry.name))
      throw new AgentDocxError("PROJECT_INVALID", "Duplicate attachment name");
    names.add(entry.name);
    const payload = attachmentPayloadPath(entry.payloadPath);
    const payloadAbsolute = resolve(directory, payload);
    const relativePayload = relative(directory, payloadAbsolute);
    if (
      relativePayload === ".." ||
      relativePayload.startsWith(`..${sep}`) ||
      isAbsolute(relativePayload)
    )
      throw new AgentDocxError(
        "PROJECT_INVALID",
        "Attachment payload escapes its bundle",
      );
    await assertRegularFile(payloadAbsolute, `Attachment ${entry.name}`);
    const bytes = await readFile(payloadAbsolute);
    if (
      bytes.byteLength !== entry.byteLength ||
      objectId(bytes) !== entry.sha256
    )
      throw new AgentDocxError(
        "PROJECT_INVALID",
        `Attachment payload hash mismatch: ${entry.name}`,
      );
  }
};
const verifyArtifactDirectory = async (
  artifactStorePath: string,
  provenanceSha256: RevisionId,
  docxSha256: RevisionId,
  attachmentManifestSha256: RevisionId | null,
  pdfSha256: RevisionId | null,
): Promise<void> => {
  const artifactDirectory = dirname(artifactStorePath);
  await assertDirectory(artifactDirectory, "Published artifact");
  if (
    (await hashRegularFile(artifactStorePath, "Published artifact DOCX")) !==
    docxSha256
  )
    throw new AgentDocxError(
      "PROJECT_INVALID",
      "Published artifact DOCX hash does not match export intent",
    );
  if (pdfSha256 !== null && pdfSha256 !== emptyObjectId) {
    const pdfPath = resolve(artifactDirectory, "document.pdf");
    if (
      (await hashRegularFile(pdfPath, "Published artifact PDF")) !== pdfSha256
    )
      throw new AgentDocxError(
        "PROJECT_INVALID",
        "Published artifact PDF hash does not match export intent",
      );
  }
  const provenancePath = resolve(artifactDirectory, "provenance.json");
  if (
    (await hashRegularFile(provenancePath, "Published artifact provenance")) !==
    provenanceSha256
  )
    throw new AgentDocxError(
      "PROJECT_INVALID",
      "Published artifact provenance hash does not match export intent",
    );
  const attachmentDirectory = resolve(artifactDirectory, "attachments");
  if (attachmentManifestSha256 === null) {
    if (await pathExists(attachmentDirectory))
      throw new AgentDocxError(
        "PROJECT_INVALID",
        "Artifact has unexpected attachments",
      );
  } else
    await verifyAttachmentDirectory(
      attachmentDirectory,
      attachmentManifestSha256,
    );
};

const publishStagedDirectory = async (
  stage: string,
  destination: string,
): Promise<void> => {
  await mkdir(destination, { recursive: false, mode: 0o700 });
  const visit = async (source: string, target: string): Promise<void> => {
    const entries = await readdir(source, { withFileTypes: true });
    for (const entry of entries.sort((left, right) =>
      left.name < right.name ? -1 : left.name > right.name ? 1 : 0,
    )) {
      if (source === stage && entry.name === "owner.json") continue;
      const sourcePath = resolve(source, entry.name);
      const targetPath = resolve(target, entry.name);
      if (entry.isSymbolicLink())
        throw new AgentDocxError(
          "PROJECT_INVALID",
          `Export stage contains a symlink: ${sourcePath}`,
        );
      if (entry.isDirectory()) {
        await mkdir(targetPath, { recursive: false, mode: 0o700 });
        await visit(sourcePath, targetPath);
      } else if (entry.isFile()) {
        await link(sourcePath, targetPath);
      } else {
        throw new AgentDocxError(
          "PROJECT_INVALID",
          `Export stage contains an unsupported entry: ${sourcePath}`,
        );
      }
    }
  };
  try {
    await visit(stage, destination);
  } catch (error) {
    await rm(destination, { recursive: true, force: true });
    throw error;
  }
};

export const recoverExport = async (
  projectDirectory: string,
  manifestPath: string,
): Promise<void> => {
  const intentPath = exportIntentPath(projectDirectory);
  let raw: unknown;
  try {
    raw = await readJsonFile<unknown>(intentPath);
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === "ENOENT") return;
    throw error;
  }
  const intent = exportIntentFrom(projectDirectory, raw);
  if (intent.manifestPath !== manifestPath)
    throw new AgentDocxError(
      "PROJECT_INVALID",
      "Export intent belongs to another manifest",
    );
  const manifest = validateManifest(await readJsonFile<unknown>(manifestPath));
  if (manifest.projectId !== intent.projectId)
    throw new AgentDocxError(
      "PROJECT_INVALID",
      "Export intent belongs to another project",
    );
  const stageIsOwned = await ownedStage(
    intent.stagePath,
    intent.owner,
    intent.projectId,
    intent.manifestPath,
  );
  if (intent.state === "preparing") {
    if (
      (await pathExists(intent.outputPath)) ||
      (intent.attachmentPath !== null &&
        (await pathExists(intent.attachmentPath)))
    )
      throw new AgentDocxError(
        "PROJECT_INVALID",
        "Preparing export has an unexpected published output",
      );
    if (stageIsOwned) await removeOwnedStage(intent);
    await rm(intentPath, { force: true });
    return;
  }
  const artifactDirectory = dirname(intent.artifactStorePath);
  const artifactExists = await pathExists(artifactDirectory);
  if (!artifactExists) {
    const artifactStagePath = resolve(intent.stagePath, "artifact");
    if (!stageIsOwned || !(await pathExists(artifactStagePath)))
      throw new AgentDocxError(
        "PROJECT_INVALID",
        "Prepared export is missing its artifact stage",
      );
    await assertNoSymlinkComponents(
      dirname(artifactDirectory),
      "Artifact parent",
    );
    await mkdir(dirname(artifactDirectory), { recursive: true, mode: 0o700 });
    await publishStagedDirectory(artifactStagePath, artifactDirectory);
  }
  const pdfSha256 = intent.pdfSha256 ?? emptyObjectId;
  const pdfDeclared = pdfSha256 !== emptyObjectId;
  await verifyArtifactDirectory(
    intent.artifactStorePath,
    intent.artifactProvenanceSha256,
    intent.docxSha256,
    intent.attachmentManifestSha256,
    pdfSha256,
  );
  await assertNoSymlinkComponents(
    intent.outputPath,
    pdfDeclared ? "Published PDF" : "Published DOCX",
  );
  const outputExists = await pathExists(intent.outputPath);
  const outputSha256 = pdfDeclared ? pdfSha256 : intent.docxSha256;
  const outputLabel = pdfDeclared ? "Published PDF" : "Published DOCX";
  if (outputExists) {
    if (
      (await hashRegularFile(intent.outputPath, outputLabel)) !== outputSha256
    )
      throw new AgentDocxError(
        "PROJECT_INVALID",
        `${outputLabel} hash does not match export intent`,
      );
  } else {
    if (!stageIsOwned)
      throw new AgentDocxError(
        "PROJECT_INVALID",
        "Prepared export is missing its stage",
      );
    const stageOutputPath = pdfDeclared
      ? intent.pdfStagePath
      : intent.docxStagePath;
    if (stageOutputPath === null)
      throw new AgentDocxError(
        "PROJECT_INVALID",
        "Prepared PDF export is missing its PDF stage",
      );
    if (
      pdfDeclared &&
      !(await pathExists(stageOutputPath)) &&
      intent.pdfStorePath !== null
    ) {
      await assertNoSymlinkComponents(stageOutputPath, "Staged PDF");
      await link(intent.pdfStorePath, stageOutputPath);
    }
    if (!(await pathExists(stageOutputPath)))
      throw new AgentDocxError(
        "PROJECT_INVALID",
        pdfDeclared
          ? "Prepared export is missing its PDF stage"
          : "Prepared export is missing its DOCX stage",
      );
    if (
      (await hashRegularFile(
        stageOutputPath,
        `Staged ${pdfDeclared ? "PDF" : "DOCX"}`,
      )) !== outputSha256
    )
      throw new AgentDocxError(
        "PROJECT_INVALID",
        `Staged ${pdfDeclared ? "PDF" : "DOCX"} hash does not match export intent`,
      );
    await link(stageOutputPath, intent.outputPath);
  }
  if (intent.attachmentPath !== null) {
    if (
      intent.attachmentManifestSha256 === null ||
      intent.attachmentStagePath === null
    )
      throw new AgentDocxError(
        "PROJECT_INVALID",
        "Prepared attachment export is incomplete",
      );
    if (!(await pathExists(intent.attachmentPath))) {
      if (!stageIsOwned || !(await pathExists(intent.attachmentStagePath)))
        throw new AgentDocxError(
          "PROJECT_INVALID",
          "Prepared export is missing its attachment stage",
        );
      await publishStagedDirectory(
        intent.attachmentStagePath,
        intent.attachmentPath,
      );
    }
    await verifyAttachmentDirectory(
      intent.attachmentPath,
      intent.attachmentManifestSha256,
    );
  }
  if (stageIsOwned)
    await rm(intent.stagePath, { recursive: true, force: true });
  await rm(intentPath, { force: true });
};

export const updateExportIntent = async (
  projectDirectory: string,
  intent: ExportIntent,
): Promise<void> => {
  const path = exportIntentPath(projectDirectory);
  const stage = `${path}.${intent.owner}.stage`;
  await writeExclusiveFile(stage, canonicalJson(intent));
  try {
    await rename(stage, path);
  } finally {
    await rm(stage, { force: true });
  }
};

export const clearExportIntent = async (
  projectDirectory: string,
): Promise<void> => {
  await rm(exportIntentPath(projectDirectory), { force: true });
};

export const completeExportIntent = async (
  projectDirectory: string,
  manifestPath: string,
): Promise<void> => recoverExport(projectDirectory, manifestPath);

export const initializeStore = async (
  manifestPath: string,
  manifest: AgentDocxManifest,
): Promise<OpenedStore> => {
  const absoluteManifestPath = resolve(manifestPath);
  const projectDirectory = dirname(absoluteManifestPath);
  validateManifest(manifest);
  await validateManifestPaths(projectDirectory, manifest);
  await assertDirectory(
    projectDirectory,
    "Project directory",
    projectDirectory,
  );
  const manifestName = basename(absoluteManifestPath);
  if (
    manifestName === STORE_DIR ||
    manifestName === INIT_INTENT ||
    manifestName === EXPORT_INTENT ||
    manifestName === LOCK_NAME ||
    manifestName.startsWith(".agent-docx.init-")
  )
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      "Manifest basename is reserved",
    );
  // Lazy to keep the lock->recovery edge one-way (lock.ts imports this module).
  const { acquireProjectLock } = await import("./lock.js");
  const release = await acquireProjectLock(projectDirectory);
  try {
    await recoverInitialization(projectDirectory, absoluteManifestPath);
    await sweepOwnedStages(projectDirectory, absoluteManifestPath);
    const storePath = resolve(projectDirectory, STORE_DIR);
    try {
      await lstat(absoluteManifestPath);
      throw new AgentDocxError(
        "DOCUMENT_EXISTS",
        `Manifest already exists: ${absoluteManifestPath}`,
      );
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code !== "ENOENT") throw error;
    }
    try {
      await lstat(storePath);
      throw new AgentDocxError(
        "DOCUMENT_EXISTS",
        `Project store already exists: ${storePath}`,
      );
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code !== "ENOENT") throw error;
    }
    const intentPath = resolve(projectDirectory, INIT_INTENT);
    const intent: InitializationIntent = {
      schemaVersion: 1,
      state: "preparing",
      projectId: manifest.projectId,
      manifestPath: absoluteManifestPath,
      storePath,
    };
    await writeAtomicFile(intentPath, canonicalJson(intent));
    await mkdir(storePath, { recursive: false, mode: 0o700 });
    await writeAtomicFile(
      bindingPath(storePath),
      canonicalJson({
        schemaVersion: 1,
        projectId: manifest.projectId,
        manifestBasename: manifestName,
      } satisfies ProjectBinding),
    );
    await writeAtomicFile(absoluteManifestPath, canonicalJson(manifest));
    await rm(intentPath, { force: true });
    return {
      manifestPath: absoluteManifestPath,
      projectDirectory: await realpath(projectDirectory),
      storePath,
      manifest,
    };
  } finally {
    await release();
  }
};
export const removeInitializedProject = async (
  manifestPath: string,
  manifest: AgentDocxManifest,
): Promise<void> => {
  const absoluteManifestPath = resolve(manifestPath);
  const projectDirectory = dirname(absoluteManifestPath);
  const storePath = resolve(projectDirectory, STORE_DIR);
  try {
    await assertRegularFile(absoluteManifestPath, "Project manifest");
    if (
      new TextDecoder("utf-8", { fatal: true }).decode(
        await readFile(absoluteManifestPath),
      ) !== canonicalJson(manifest)
    )
      return;
    const storeEntry = await lstat(storePath);
    if (!storeEntry.isDirectory() || storeEntry.isSymbolicLink()) return;
    const bindingBytes = await readFile(bindingPath(storePath));
    const binding = {
      schemaVersion: 1,
      projectId: manifest.projectId,
      manifestBasename: basename(absoluteManifestPath),
    } satisfies ProjectBinding;
    if (
      new TextDecoder("utf-8", { fatal: true }).decode(bindingBytes) !==
      canonicalJson(binding)
    )
      return;
    const entries = await readdir(storePath, { withFileTypes: true });
    if (
      entries.length !== 1 ||
      entries[0]!.name !== BINDING_FILE ||
      !entries[0]!.isFile() ||
      entries[0]!.isSymbolicLink()
    )
      return;
    await rm(storePath, { recursive: true, force: false });
    await rm(absoluteManifestPath, { force: false });
  } catch {
    return;
  }
};

export const documentConfigFromInput = async (
  projectDirectory: string,
  input: ProjectDocumentInput,
): Promise<AgentDocxDocumentConfig> => {
  if (!isDocumentId(input.documentId))
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      "document id must be a lowercase slug",
    );
  const normalize = async (path: string, name: string): Promise<string> => {
    const absolutePath = assertWithin(projectDirectory, path, name);
    await assertRegularFile(absolutePath, name, projectDirectory);
    return assertRelativeManifestPath(
      relativePath(projectDirectory, absolutePath),
      name,
    );
  };
  const sourcePath = assertWithin(
    projectDirectory,
    input.source,
    "Document source",
  );
  if (input.createSource) {
    try {
      await lstat(sourcePath);
      throw new AgentDocxError(
        "OUTPUT_EXISTS",
        `Source already exists: ${input.source}`,
      );
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code !== "ENOENT") throw error;
    }
  } else {
    await assertRegularFile(sourcePath, "Document source", projectDirectory);
  }
  const source = assertRelativeManifestPath(
    relativePath(projectDirectory, sourcePath),
    "Document source",
  );
  const fontSet = input.fontSet
    ? {
        family: input.fontSet.family,
        regularPath: await normalize(input.fontSet.regularPath, "Font regular"),
        ...(input.fontSet.boldPath
          ? { boldPath: await normalize(input.fontSet.boldPath, "Font bold") }
          : {}),
        ...(input.fontSet.italicPath
          ? {
              italicPath: await normalize(
                input.fontSet.italicPath,
                "Font italic",
              ),
            }
          : {}),
        ...(input.fontSet.boldItalicPath
          ? {
              boldItalicPath: await normalize(
                input.fontSet.boldItalicPath,
                "Font boldItalic",
              ),
            }
          : {}),
      }
    : undefined;
  if (
    (input.fontSet?.family === undefined) !==
    (input.fontSet?.regularPath === undefined)
  )
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      "Font family and regular font are paired",
    );
  let assetsDir: string | undefined;
  if (input.assetsDir) {
    const absoluteAssets = assertWithin(
      projectDirectory,
      input.assetsDir,
      "Assets directory",
    );
    await assertDirectory(absoluteAssets, "Assets directory", projectDirectory);
    assetsDir = assertRelativeManifestPath(
      relativePath(projectDirectory, absoluteAssets),
      "Assets directory",
    );
  }
  let rulePacks: readonly string[] | undefined;
  if (input.rulePacks !== undefined) {
    if (!Array.isArray(input.rulePacks))
      throw new AgentDocxError(
        "INVALID_ARGUMENT",
        "rulePacks must be an array",
      );
    rulePacks = await Promise.all(
      input.rulePacks.map((path, index) =>
        normalize(path, `Rule pack ${index}`),
      ),
    );
  }
  return {
    id: input.documentId,
    source,
    profile: input.profile,
    ...(input.filingKind ? { filingKind: input.filingKind } : {}),
    ...(rulePacks ? { rulePacks } : {}),
    ...(input.rulePack ? { rulePack: input.rulePack } : {}),
    ...(input.template
      ? { template: await normalize(input.template, "Template") }
      : {}),
    ...(assetsDir ? { assetsDir } : {}),
    ...(fontSet ? { fontSet } : {}),
    ...(input.chrome ? { chrome: input.chrome } : {}),
    metadata: input.metadata,
  };
};

export const createEmptySource = async (
  projectDirectory: string,
  source: string,
): Promise<void> => {
  const sourcePath = assertWithin(projectDirectory, source, "Document source");
  await assertNoSymlinkComponents(sourcePath, "Document source");
  await mkdir(dirname(sourcePath), { recursive: true, mode: 0o700 });
  await writeExclusiveFile(sourcePath, "");
};

export const replaceOwnedFile = async (
  path: string,
  expectedOld: RevisionId | null,
  bytes: Uint8Array | string,
): Promise<RevisionId> => {
  const nextId = objectId(bytes);
  await assertNoSymlinkComponents(path, "Owned file");
  if (expectedOld === null) {
    try {
      await lstat(path);
      throw new AgentDocxError(
        "WORKING_COPY_CONFLICT",
        `Refusing to replace unowned file: ${path}`,
      );
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code !== "ENOENT") throw error;
    }
  } else {
    await assertRegularFile(path, "Owned file");
    if (objectId(await readProjectFile(path, "Owned file")) !== expectedOld)
      throw new AgentDocxError(
        "WORKING_COPY_CONFLICT",
        `Owned file changed: ${path}`,
      );
  }
  const stage = `${path}.${randomUUID()}.stage`;
  const backup = `${path}.${randomUUID()}.backup`;
  await writeExclusiveFile(stage, bytes);
  let backupCreated = false;
  let replacementLinked = false;
  const rollback = async (): Promise<void> => {
    if (expectedOld === null) {
      if (!replacementLinked) return;
      let currentId: RevisionId | null = null;
      try {
        const entry = await lstat(path);
        if (entry.isSymbolicLink() || !entry.isFile())
          throw new AgentDocxError(
            "PROJECT_INVALID",
            `Cannot roll back changed owned file: ${path}`,
          );
        currentId = objectId(await readProjectFile(path, "Owned file"));
      } catch (error) {
        const code =
          error instanceof AgentDocxError
            ? error.code
            : (error as NodeJS.ErrnoException).code;
        if (code === "INPUT_NOT_FOUND" || code === "ENOENT") return;
        throw error;
      }
      if (currentId !== nextId)
        throw new AgentDocxError(
          "PROJECT_INVALID",
          `Cannot roll back changed owned file: ${path}`,
        );
      await rm(path);
      return;
    }
    if (!backupCreated) return;
    let currentId: RevisionId | null = null;
    try {
      const entry = await lstat(path);
      if (entry.isSymbolicLink() || !entry.isFile())
        throw new AgentDocxError(
          "PROJECT_INVALID",
          `Cannot roll back changed owned file: ${path}`,
        );
      currentId = objectId(await readProjectFile(path, "Owned file"));
    } catch (error) {
      const code =
        error instanceof AgentDocxError
          ? error.code
          : (error as NodeJS.ErrnoException).code;
      if (code !== "INPUT_NOT_FOUND" && code !== "ENOENT") throw error;
    }
    if (currentId !== null && currentId !== nextId)
      throw new AgentDocxError(
        "PROJECT_INVALID",
        `Cannot roll back changed owned file: ${path}`,
      );
    if (currentId !== null) await rm(path);
    await rename(backup, path);
    backupCreated = false;
    await assertRegularFile(path, "Owned rollback");
    if (objectId(await readProjectFile(path, "Owned rollback")) !== expectedOld)
      throw new AgentDocxError(
        "PROJECT_INVALID",
        `Owned rollback mismatch: ${path}`,
      );
  };
  try {
    if (expectedOld === null) {
      await link(stage, path);
      replacementLinked = true;
    } else {
      await rename(path, backup);
      backupCreated = true;
      await assertRegularFile(backup, "Owned backup");
      if (
        objectId(await readProjectFile(backup, "Owned backup")) !== expectedOld
      )
        throw new AgentDocxError(
          "PROJECT_INVALID",
          `Owned backup mismatch: ${path}`,
        );
      await link(stage, path);
      replacementLinked = true;
    }
    await assertRegularFile(path, "Owned replacement");
    if (objectId(await readProjectFile(path, "Owned replacement")) !== nextId)
      throw new AgentDocxError(
        "PROJECT_INVALID",
        `Owned replacement mismatch: ${path}`,
      );
    if (expectedOld !== null) {
      await rm(backup);
      backupCreated = false;
    }
    return nextId;
  } catch (error) {
    await rollback();
    throw error;
  } finally {
    await rm(stage, { force: true });
  }
};

export const removeOwnedFile = async (
  path: string,
  expected: RevisionId,
): Promise<void> => {
  await assertRegularFile(path, "Owned file");
  if (objectId(await readProjectFile(path, "Owned file")) !== expected)
    throw new AgentDocxError(
      "WORKING_COPY_CONFLICT",
      `Owned file changed: ${path}`,
    );
  const backup = `${path}.${randomUUID()}.backup`;
  await rename(path, backup);
  try {
    await assertRegularFile(backup, "Owned backup");
    if (objectId(await readProjectFile(backup, "Owned backup")) !== expected)
      throw new AgentDocxError(
        "PROJECT_INVALID",
        `Owned backup mismatch: ${path}`,
      );
    await rm(backup);
  } catch (error) {
    try {
      await lstat(path);
      throw new AgentDocxError(
        "PROJECT_INVALID",
        `Cannot roll back removed owned file: ${path}`,
      );
    } catch (restoreError) {
      if ((restoreError as NodeJS.ErrnoException).code !== "ENOENT")
        throw restoreError;
    }
    try {
      await rename(backup, path);
      await assertRegularFile(path, "Owned rollback");
      if (objectId(await readProjectFile(path, "Owned rollback")) !== expected)
        throw new AgentDocxError(
          "PROJECT_INVALID",
          `Owned rollback mismatch: ${path}`,
        );
    } catch (restoreError) {
      throw new AgentDocxError(
        "PROJECT_INVALID",
        `Cannot restore removed owned file: ${path}`,
        {
          cause:
            restoreError instanceof Error
              ? restoreError.message
              : String(restoreError),
        },
      );
    }
    throw error;
  }
};

export const updateManifest = async (
  opened: OpenedStore,
  manifest: AgentDocxManifest,
): Promise<OpenedStore> => {
  validateManifest(manifest);
  await validateManifestPaths(opened.projectDirectory, manifest);
  const stage = `${opened.manifestPath}.${randomUUID()}.stage`;
  await writeExclusiveFile(stage, canonicalJson(manifest));
  try {
    await rename(stage, opened.manifestPath);
  } finally {
    await rm(stage, { force: true });
  }
  return { ...opened, manifest };
};
