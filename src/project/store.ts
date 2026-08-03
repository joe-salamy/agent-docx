import { createHash, randomUUID } from "node:crypto";
import {
  link,
  lstat,
  mkdir,
  open,
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
  parse,
  relative,
  resolve,
  sep,
} from "node:path";
import canonicalize from "canonicalize";
import { lock } from "proper-lockfile";
import { AgentDocxError } from "../types.js";
import { builtInProfiles } from "../profiles.js";
import { isDocumentId, type RevisionId } from "../legal/model.js";
import { builtInRulePacks } from "../legal/rules.js";
import type {
  AgentDocxDocumentConfig,
  AgentDocxManifest,
  DependencyHashes,
  ProjectDocumentInput,
} from "./contracts.js";

export const STORE_DIR = ".agent-docx";
const INIT_INTENT = ".agent-docx.init.json";
export const EXPORT_INTENT = ".agent-docx.export.json";
const LOCK_NAME = ".agent-docx.lock";
const BINDING_FILE = "project.json";

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

export type ExportIntent = {
  schemaVersion: 1;
  state: "preparing" | "prepared";
  projectId: string;
  manifestPath: string;
  owner: string;
  outputPath: string;
  attachmentPath: string | null;
  stagePath: string;
  docxStagePath: string;
  attachmentStagePath: string | null;
  docxSha256: RevisionId;
  attachmentManifestSha256: RevisionId | null;
  artifactProvenanceSha256: RevisionId;
  artifactStorePath: string;
  attachmentStorePath: string | null;
};
export type OpenedStore = {
  manifestPath: string;
  projectDirectory: string;
  storePath: string;
  manifest: AgentDocxManifest;
};

export type ProjectSnapshot = {
  source: string;
  sourceObject: RevisionId;
  documentConfigObject: RevisionId;
  dependencyObjects: DependencyHashes;
  dependencyBytes: ReadonlyMap<
    string,
    { bytes: Uint8Array; mediaType: string }
  >;
  workingTreeHash: RevisionId;
};

const allowedManifestKeys: Record<string, true> = {
  schemaVersion: true,
  projectId: true,
  defaultDocument: true,
  storeDir: true,
  documents: true,
};

const allowedDocumentKeys: Record<string, true> = {
  id: true,
  source: true,
  profile: true,
  filingKind: true,
  rulePack: true,
  template: true,
  assetsDir: true,
  fontSet: true,
  chrome: true,
  metadata: true,
};

const fileMode = 0o600;

export const objectId = (bytes: Uint8Array | string): RevisionId =>
  `sha256:${createHash("sha256").update(bytes).digest("hex")}`;

export const canonicalJson = (value: unknown): string => {
  const serialized = canonicalize(value);
  if (serialized === undefined)
    throw new AgentDocxError(
      "PROJECT_INVALID",
      "Value cannot be canonicalized",
    );
  return serialized;
};

export const canonicalObjectId = (value: unknown): RevisionId =>
  objectId(canonicalJson(value));

const relativePath = (projectDirectory: string, absolutePath: string): string =>
  relative(projectDirectory, absolutePath).split(sep).join("/");

const assertRelativeManifestPath = (path: string, name: string): string => {
  const parts = path.split("/");
  const reservedRoot =
    parts[0] === STORE_DIR ||
    parts[0] === INIT_INTENT ||
    parts[0] === EXPORT_INTENT ||
    parts[0] === LOCK_NAME ||
    parts[0]?.startsWith(".agent-docx.init-") === true;
  if (
    path.length === 0 ||
    isAbsolute(path) ||
    path.includes("\\") ||
    reservedRoot ||
    parts.some((part) => part.length === 0 || part === "." || part === "..")
  )
    throw new AgentDocxError(
      "PATH_OUTSIDE_PROJECT",
      `${name} must be a normalized project-relative path`,
    );
  return path;
};

export const assertNoSymlinkComponents = async (
  path: string,
  name: string,
): Promise<void> => {
  const absolutePath = resolve(path);
  const root = parse(absolutePath).root;
  let current = root;
  for (const component of relative(root, absolutePath).split(sep)) {
    if (component.length === 0) continue;
    current = resolve(current, component);
    let entry;
    try {
      entry = await lstat(current);
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code === "ENOENT") return;
      throw error;
    }
    if (entry.isSymbolicLink())
      throw new AgentDocxError(
        "PATH_OUTSIDE_PROJECT",
        `${name} contains a symbolic-link component`,
      );
  }
};

const assertRegularFile = async (path: string, name: string): Promise<void> => {
  await assertNoSymlinkComponents(path, name);
  let entry;
  try {
    entry = await lstat(path);
  } catch {
    throw new AgentDocxError(
      "INPUT_NOT_FOUND",
      `${name} does not exist: ${path}`,
    );
  }
  if (!entry.isFile() || entry.isSymbolicLink())
    throw new AgentDocxError(
      "PATH_OUTSIDE_PROJECT",
      `${name} must be a regular nonsymlink file`,
    );
};

const assertDirectory = async (path: string, name: string): Promise<void> => {
  await assertNoSymlinkComponents(path, name);
  let entry;
  try {
    entry = await lstat(path);
  } catch {
    throw new AgentDocxError(
      "INPUT_NOT_FOUND",
      `${name} does not exist: ${path}`,
    );
  }
  if (!entry.isDirectory() || entry.isSymbolicLink())
    throw new AgentDocxError(
      "PATH_OUTSIDE_PROJECT",
      `${name} must be a directory, not a symlink`,
    );
};

const assertWithin = (
  projectDirectory: string,
  path: string,
  name: string,
): string => {
  const normalizedPath = assertRelativeManifestPath(path, name);
  const absolutePath = resolve(projectDirectory, normalizedPath);
  const relativePathValue = relative(projectDirectory, absolutePath);
  if (
    relativePathValue === "" ||
    relativePathValue === ".." ||
    relativePathValue.startsWith(`..${sep}`) ||
    isAbsolute(relativePathValue)
  )
    throw new AgentDocxError(
      "PATH_OUTSIDE_PROJECT",
      `${name} is outside the project directory`,
    );
  return absolutePath;
};

const strictJson = <T>(bytes: Uint8Array, path: string): T => {
  try {
    return JSON.parse(
      new TextDecoder("utf-8", { fatal: true }).decode(bytes),
    ) as T;
  } catch (error) {
    throw new AgentDocxError("PROJECT_INVALID", `Invalid JSON: ${path}`, {
      cause: error instanceof Error ? error.message : String(error),
    });
  }
};

const hasOnlyKeys = (
  value: Record<string, unknown>,
  allowed: Record<string, true>,
): boolean => Object.keys(value).every((key) => allowed[key] === true);

const objectRecord = (
  value: unknown,
  name: string,
): Record<string, unknown> => {
  if (!value || typeof value !== "object" || Array.isArray(value))
    throw new AgentDocxError("PROJECT_INVALID", `${name} must be an object`);
  return value as Record<string, unknown>;
};

const assertClosedKeys = (
  value: Record<string, unknown>,
  keys: readonly string[],
  name: string,
): void => {
  if (!Object.keys(value).every((key) => keys.includes(key)))
    throw new AgentDocxError(
      "PROJECT_INVALID",
      `${name} has unknown properties`,
    );
};

const assertString = (value: unknown, name: string): string => {
  if (typeof value !== "string")
    throw new AgentDocxError("PROJECT_INVALID", `${name} must be a string`);
  return value;
};

const assertMetadata = (value: unknown): void => {
  const metadata = objectRecord(value, "Document metadata");
  assertClosedKeys(
    metadata,
    [
      "court",
      "jurisdiction",
      "caseName",
      "docketNumber",
      "documentTitle",
      "filingDate",
      "parties",
      "counsel",
      "certificates",
    ],
    "Document metadata",
  );
  for (const key of [
    "court",
    "jurisdiction",
    "caseName",
    "docketNumber",
    "documentTitle",
  ] as const)
    assertString(metadata[key], `Document metadata ${key}`);
  if (
    metadata.filingDate !== undefined &&
    (typeof metadata.filingDate !== "string" ||
      !/^\d{4}-\d{2}-\d{2}$/.test(metadata.filingDate))
  )
    throw new AgentDocxError(
      "PROJECT_INVALID",
      "Document metadata filingDate is invalid",
    );
  for (const [key, allowed] of [
    ["parties", ["id", "name", "role"]],
    [
      "counsel",
      ["id", "name", "barNumber", "firm", "addressLines", "phone", "email"],
    ],
  ] as const) {
    if (!Array.isArray(metadata[key]))
      throw new AgentDocxError(
        "PROJECT_INVALID",
        `Document metadata ${key} must be an array`,
      );
    const ids = new Set<string>();
    for (const entry of metadata[key]) {
      const record = objectRecord(entry, `Document metadata ${key} entry`);
      assertClosedKeys(record, allowed, `Document metadata ${key} entry`);
      const id = assertString(record.id, `Document metadata ${key} entry id`);
      if (!isDocumentId(id) || ids.has(id))
        throw new AgentDocxError(
          "PROJECT_INVALID",
          `Document metadata ${key} entry id is invalid`,
        );
      ids.add(id);
      assertString(record.name, `Document metadata ${key} entry name`);
      if (key === "parties")
        assertString(record.role, "Document metadata party role");
      if (record.addressLines !== undefined) {
        if (
          !Array.isArray(record.addressLines) ||
          record.addressLines.some((line) => typeof line !== "string")
        )
          throw new AgentDocxError(
            "PROJECT_INVALID",
            "Document metadata counsel addressLines is invalid",
          );
      }
      for (const optional of ["barNumber", "firm", "phone", "email"] as const)
        if (record[optional] !== undefined)
          assertString(
            record[optional],
            `Document metadata counsel ${optional}`,
          );
    }
  }
  if (!Array.isArray(metadata.certificates))
    throw new AgentDocxError(
      "PROJECT_INVALID",
      "Document metadata certificates must be an array",
    );
  const certificateIds = new Set<string>();
  for (const certificate of metadata.certificates) {
    const record = objectRecord(certificate, "Document metadata certificate");
    const id = assertString(record.id, "Document metadata certificate id");
    if (!isDocumentId(id) || certificateIds.has(id))
      throw new AgentDocxError(
        "PROJECT_INVALID",
        "Document metadata certificate id is invalid",
      );
    certificateIds.add(id);
    if (record.kind === "service") {
      assertClosedKeys(
        record,
        [
          "id",
          "kind",
          "statement",
          "servedOn",
          "method",
          "date",
          "signerCounselId",
        ],
        "Service certificate",
      );
      assertString(record.statement, "Service certificate statement");
      assertString(record.method, "Service certificate method");
      assertString(
        record.signerCounselId,
        "Service certificate signerCounselId",
      );
      if (
        !Array.isArray(record.servedOn) ||
        record.servedOn.some((entry) => typeof entry !== "string")
      )
        throw new AgentDocxError(
          "PROJECT_INVALID",
          "Service certificate servedOn is invalid",
        );
      if (record.date !== undefined)
        assertString(record.date, "Service certificate date");
    } else if (record.kind === "compliance") {
      assertClosedKeys(
        record,
        ["id", "kind", "basis", "signerCounselId"],
        "Compliance certificate",
      );
      if (record.basis !== "words" && record.basis !== "monospaced-lines")
        throw new AgentDocxError(
          "PROJECT_INVALID",
          "Compliance certificate basis is invalid",
        );
      assertString(
        record.signerCounselId,
        "Compliance certificate signerCounselId",
      );
    } else
      throw new AgentDocxError(
        "PROJECT_INVALID",
        "Certificate kind is invalid",
      );
  }
};

const assertChrome = (value: unknown): void => {
  const chrome = objectRecord(value, "Document chrome");
  assertClosedKeys(
    chrome,
    ["headers", "footers", "pageNumber", "lineNumbers"],
    "Document chrome",
  );
  for (const storyKind of ["headers", "footers"] as const) {
    if (chrome[storyKind] === undefined) continue;
    const stories = objectRecord(
      chrome[storyKind],
      `Document chrome ${storyKind}`,
    );
    assertClosedKeys(
      stories,
      ["default", "first", "even"],
      `Document chrome ${storyKind}`,
    );
    for (const text of Object.values(stories)) {
      const template = assertString(text, `Document chrome ${storyKind} value`);
      for (const token of template.matchAll(/\{\{([^}]+)\}\}/g))
        if (
          ![
            "caseName",
            "docketNumber",
            "documentTitle",
            "page",
            "pages",
          ].includes(token[1]!)
        )
          throw new AgentDocxError(
            "PROJECT_INVALID",
            `Unknown document chrome token: ${token[1]!}`,
          );
    }
  }
  if (chrome.pageNumber !== undefined) {
    const number = objectRecord(
      chrome.pageNumber,
      "Document chrome pageNumber",
    );
    assertClosedKeys(
      number,
      ["story", "alignment", "format", "start"],
      "Document chrome pageNumber",
    );
    if (number.story !== "header" && number.story !== "footer")
      throw new AgentDocxError(
        "PROJECT_INVALID",
        "Document chrome pageNumber story is invalid",
      );
    if (!["left", "center", "right"].includes(String(number.alignment)))
      throw new AgentDocxError(
        "PROJECT_INVALID",
        "Document chrome pageNumber alignment is invalid",
      );
    if (
      !["decimal", "lower-roman", "upper-roman"].includes(String(number.format))
    )
      throw new AgentDocxError(
        "PROJECT_INVALID",
        "Document chrome pageNumber format is invalid",
      );
    if (!Number.isInteger(number.start) || (number.start as number) < 1)
      throw new AgentDocxError(
        "PROJECT_INVALID",
        "Document chrome pageNumber start is invalid",
      );
  }
  if (chrome.lineNumbers !== undefined) {
    const lines = objectRecord(
      chrome.lineNumbers,
      "Document chrome lineNumbers",
    );
    assertClosedKeys(
      lines,
      ["countBy", "start", "distanceTwips", "restart"],
      "Document chrome lineNumbers",
    );
    for (const numeric of ["countBy", "start", "distanceTwips"] as const)
      if (
        !Number.isInteger(lines[numeric]) ||
        (lines[numeric] as number) < (numeric === "distanceTwips" ? 0 : 1)
      )
        throw new AgentDocxError(
          "PROJECT_INVALID",
          `Document chrome lineNumbers ${numeric} is invalid`,
        );
    if (
      !["continuous", "new-page", "new-section"].includes(String(lines.restart))
    )
      throw new AgentDocxError(
        "PROJECT_INVALID",
        "Document chrome lineNumbers restart is invalid",
      );
  }
};

const validateDocumentConfig = (value: unknown): AgentDocxDocumentConfig => {
  if (!value || typeof value !== "object" || Array.isArray(value))
    throw new AgentDocxError(
      "PROJECT_INVALID",
      "Document config must be an object",
    );
  const config = value as Record<string, unknown>;
  if (!hasOnlyKeys(config, allowedDocumentKeys))
    throw new AgentDocxError(
      "PROJECT_INVALID",
      "Document config has unknown properties",
    );
  if (typeof config.id !== "string" || !isDocumentId(config.id))
    throw new AgentDocxError("PROJECT_INVALID", "Document id is invalid");
  if (typeof config.source !== "string")
    throw new AgentDocxError("PROJECT_INVALID", "Document source is invalid");
  assertRelativeManifestPath(config.source, "Document source");
  if (
    typeof config.profile !== "string" ||
    !(config.profile in builtInProfiles)
  )
    throw new AgentDocxError("PROJECT_INVALID", "Document profile is invalid");
  assertMetadata(config.metadata);
  if (
    config.filingKind !== undefined &&
    ![
      "principal-brief",
      "reply-brief",
      "motion-document",
      "opposition-text",
      "reply-text",
    ].includes(String(config.filingKind))
  )
    throw new AgentDocxError(
      "PROJECT_INVALID",
      "Document filingKind is invalid",
    );
  if (
    config.rulePack !== undefined &&
    !["frap-32@2024-12-01", "cand-civil@2026-05-01"].includes(
      String(config.rulePack),
    )
  )
    throw new AgentDocxError("PROJECT_INVALID", "Document rulePack is invalid");
  if (config.chrome !== undefined) assertChrome(config.chrome);
  for (const name of ["template", "assetsDir"] as const) {
    const path = config[name];
    if (path !== undefined) {
      if (typeof path !== "string")
        throw new AgentDocxError("PROJECT_INVALID", `${name} is invalid`);
      assertRelativeManifestPath(path, name);
    }
  }
  if (config.fontSet !== undefined) {
    if (
      !config.fontSet ||
      typeof config.fontSet !== "object" ||
      Array.isArray(config.fontSet)
    )
      throw new AgentDocxError("PROJECT_INVALID", "fontSet is invalid");
    const fontSet = objectRecord(config.fontSet, "fontSet");
    assertClosedKeys(
      fontSet,
      ["family", "regularPath", "boldPath", "italicPath", "boldItalicPath"],
      "fontSet",
    );
    if (typeof fontSet.family !== "string" || fontSet.family.length === 0)
      throw new AgentDocxError("PROJECT_INVALID", "fontSet family is invalid");
    if (typeof fontSet.regularPath !== "string")
      throw new AgentDocxError(
        "PROJECT_INVALID",
        "fontSet regularPath is invalid",
      );
    for (const role of [
      "regularPath",
      "boldPath",
      "italicPath",
      "boldItalicPath",
    ] as const) {
      const path = fontSet[role];
      if (path !== undefined) {
        if (typeof path !== "string")
          throw new AgentDocxError(
            "PROJECT_INVALID",
            `fontSet ${role} is invalid`,
          );
        assertRelativeManifestPath(path, `fontSet ${role}`);
      }
    }
  }
  return config as unknown as AgentDocxDocumentConfig;
};

export const validateManifest = (value: unknown): AgentDocxManifest => {
  if (!value || typeof value !== "object" || Array.isArray(value))
    throw new AgentDocxError(
      "PROJECT_INVALID",
      "Project manifest must be an object",
    );
  const manifest = value as Record<string, unknown>;
  if (!hasOnlyKeys(manifest, allowedManifestKeys))
    throw new AgentDocxError(
      "PROJECT_INVALID",
      "Project manifest has unknown properties",
    );
  if (manifest.schemaVersion !== 1)
    throw new AgentDocxError(
      "PROJECT_INVALID",
      "Project manifest schemaVersion must be 1",
    );
  if (
    typeof manifest.projectId !== "string" ||
    !/^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/.test(
      manifest.projectId,
    )
  )
    throw new AgentDocxError("PROJECT_INVALID", "Project id is invalid");
  if (
    typeof manifest.defaultDocument !== "string" ||
    !isDocumentId(manifest.defaultDocument)
  )
    throw new AgentDocxError(
      "PROJECT_INVALID",
      "Default document id is invalid",
    );
  if (manifest.storeDir !== STORE_DIR)
    throw new AgentDocxError(
      "PROJECT_INVALID",
      `Project storeDir must be ${STORE_DIR}`,
    );
  if (!Array.isArray(manifest.documents) || manifest.documents.length === 0)
    throw new AgentDocxError(
      "PROJECT_INVALID",
      "Project must contain at least one document",
    );
  const documents = manifest.documents.map(validateDocumentConfig);
  const ids = new Set<string>();
  const paths = new Set<string>();
  for (const document of documents) {
    if (ids.has(document.id))
      throw new AgentDocxError(
        "PROJECT_INVALID",
        `Duplicate document id: ${document.id}`,
      );
    if (paths.has(document.source))
      throw new AgentDocxError(
        "PROJECT_INVALID",
        `Duplicate document source: ${document.source}`,
      );
    ids.add(document.id);
    paths.add(document.source);
  }
  if (!ids.has(manifest.defaultDocument))
    throw new AgentDocxError(
      "PROJECT_INVALID",
      "Default document is not declared",
    );
  return {
    schemaVersion: 1,
    projectId: manifest.projectId,
    defaultDocument: manifest.defaultDocument,
    storeDir: STORE_DIR,
    documents,
  };
};
type OwnedManifestPath = {
  path: string;
  documentId: string;
  shareableFont: boolean;
};

export const validateManifestPaths = async (
  projectDirectory: string,
  manifest: AgentDocxManifest,
): Promise<void> => {
  const owned: OwnedManifestPath[] = [];
  const add = (
    path: string | undefined,
    label: string,
    documentId: string,
    shareableFont = false,
  ): void => {
    if (path === undefined) return;
    const normalized = assertRelativeManifestPath(path, label);
    const absolute = assertWithin(projectDirectory, normalized, label);
    for (const existing of owned) {
      const sharedFont =
        absolute === existing.path && shareableFont && existing.shareableFont;
      if (pathsOverlap(absolute, existing.path) && !sharedFont)
        throw new AgentDocxError(
          "PROJECT_INVALID",
          `Manifest path is already owned: ${path}`,
        );
    }
    owned.push({ path: absolute, documentId, shareableFont });
  };
  for (const document of manifest.documents) {
    add(document.source, "Document source", document.id);
    add(document.template, "Document template", document.id);
    add(document.assetsDir, "Assets directory", document.id);
    for (const [role, path] of [
      ["regular", document.fontSet?.regularPath],
      ["bold", document.fontSet?.boldPath],
      ["italic", document.fontSet?.italicPath],
      ["boldItalic", document.fontSet?.boldItalicPath],
    ] as const)
      add(path, `Font ${role}`, document.id, true);
  }
};

const pathsOverlap = (left: string, right: string): boolean =>
  left === right ||
  left.startsWith(`${right}${sep}`) ||
  right.startsWith(`${left}${sep}`);

const readJsonFile = async <T>(path: string): Promise<T> =>
  strictJson<T>(await readFile(path), path);

const writeExclusive = async (
  path: string,
  bytes: Uint8Array | string,
): Promise<void> => {
  const handle = await open(path, "wx", fileMode);
  try {
    await handle.writeFile(bytes);
  } finally {
    await handle.close();
  }
};

const objectPath = (storePath: string, id: RevisionId): string => {
  const hex = id.slice("sha256:".length);
  if (!/^[0-9a-f]{64}$/.test(hex))
    throw new AgentDocxError("PROJECT_INVALID", `Invalid object id: ${id}`);
  return resolve(storePath, "objects", "sha256", hex.slice(0, 2), hex.slice(2));
};

export const writeObject = async (
  storePath: string,
  bytes: Uint8Array | string,
): Promise<RevisionId> => {
  const id = objectId(bytes);
  const path = objectPath(storePath, id);
  await mkdir(dirname(path), { recursive: true, mode: 0o700 });
  try {
    await writeExclusive(path, bytes);
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code !== "EEXIST") throw error;
    const current = await readFile(path);
    if (objectId(current) !== id)
      throw new AgentDocxError("PROJECT_INVALID", `Object collision at ${id}`);
  }
  return id;
};

export const readObject = async (
  storePath: string,
  id: RevisionId,
): Promise<Uint8Array> => {
  const path = objectPath(storePath, id);
  await assertRegularFile(path, "Object");
  const bytes = await readFile(path);
  if (objectId(bytes) !== id)
    throw new AgentDocxError("PROJECT_INVALID", `Object hash mismatch: ${id}`);
  return bytes;
};

const refsPath = (storePath: string, documentId: string): string => {
  if (!isDocumentId(documentId))
    throw new AgentDocxError(
      "DOCUMENT_NOT_FOUND",
      `Invalid document id: ${documentId}`,
    );
  return resolve(storePath, "refs", `${documentId}.json`);
};

const revisionPath = (storePath: string, id: RevisionId): string => {
  const hex = id.slice("sha256:".length);
  if (!/^[0-9a-f]{64}$/.test(hex))
    throw new AgentDocxError(
      "REVISION_NOT_FOUND",
      `Invalid revision id: ${id}`,
    );
  return resolve(storePath, "revisions", `${hex}.json`);
};

export const readHead = async (
  storePath: string,
  documentId: string,
): Promise<RevisionId | null> => {
  const path = refsPath(storePath, documentId);
  try {
    const entry = await readJsonFile<{
      schemaVersion: 1;
      documentId: string;
      head: RevisionId | null;
    }>(path);
    if (entry.schemaVersion !== 1 || entry.documentId !== documentId)
      throw new AgentDocxError(
        "PROJECT_INVALID",
        `Malformed head reference: ${documentId}`,
      );
    if (entry.head !== null && !/^sha256:[0-9a-f]{64}$/.test(entry.head))
      throw new AgentDocxError(
        "PROJECT_INVALID",
        `Malformed head reference: ${documentId}`,
      );
    return entry.head;
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === "ENOENT") return null;
    throw error;
  }
};

export const writeHead = async (
  storePath: string,
  documentId: string,
  head: RevisionId,
): Promise<void> => {
  const path = refsPath(storePath, documentId);
  await mkdir(dirname(path), { recursive: true, mode: 0o700 });
  const stage = `${path}.${randomUUID()}.stage`;
  await writeExclusive(
    stage,
    canonicalJson({ schemaVersion: 1, documentId, head }),
  );
  try {
    await rename(stage, path);
  } finally {
    await rm(stage, { force: true });
  }
};

export const readRevisionJson = async <T>(
  storePath: string,
  revision: RevisionId,
): Promise<T> => {
  const path = revisionPath(storePath, revision);
  const record = strictJson<Record<string, unknown>>(
    await readFile(path),
    path,
  );
  const withoutId = { ...record };
  delete withoutId.id;
  if (record.id !== revision || canonicalObjectId(withoutId) !== revision)
    throw new AgentDocxError(
      "PROJECT_INVALID",
      `Revision hash mismatch: ${revision}`,
    );
  return record as T;
};

export const writeRevisionJson = async (
  storePath: string,
  value: Record<string, unknown>,
): Promise<RevisionId> => {
  const withoutId = { ...value };
  delete withoutId.id;
  const id = canonicalObjectId(withoutId);
  const record = canonicalJson({ ...withoutId, id });
  const path = revisionPath(storePath, id);
  await mkdir(dirname(path), { recursive: true, mode: 0o700 });
  try {
    await writeExclusive(path, record);
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code !== "EEXIST") throw error;
    const existing = strictJson<Record<string, unknown>>(
      await readFile(path),
      path,
    );
    const existingWithoutId = { ...existing };
    delete existingWithoutId.id;
    if (existing.id !== id || canonicalObjectId(existingWithoutId) !== id)
      throw new AgentDocxError(
        "PROJECT_INVALID",
        `Revision collision at ${id}`,
      );
  }
  return id;
};

const mediaTypeFor = (path: string): string => {
  const lower = path.toLowerCase();
  if (lower.endsWith(".png")) return "image/png";
  if (lower.endsWith(".jpg") || lower.endsWith(".jpeg")) return "image/jpeg";
  if (lower.endsWith(".pdf")) return "application/pdf";
  return "application/octet-stream";
};

const collectAssets = async (
  directory: string,
  relativeDirectory = "",
): Promise<ReadonlyMap<string, { bytes: Uint8Array; mediaType: string }>> => {
  const assets = new Map<string, { bytes: Uint8Array; mediaType: string }>();
  const entries = await readdir(directory, { withFileTypes: true });
  for (const entry of entries.sort((left, right) =>
    left.name.localeCompare(right.name),
  )) {
    const logical = relativeDirectory
      ? `${relativeDirectory}/${entry.name}`
      : entry.name;
    const path = resolve(directory, entry.name);
    if (entry.isSymbolicLink())
      throw new AgentDocxError(
        "PATH_OUTSIDE_PROJECT",
        `Asset is a symlink: ${logical}`,
      );
    if (entry.isDirectory()) {
      for (const [key, value] of await collectAssets(path, logical))
        assets.set(key, value);
      continue;
    }
    if (!entry.isFile())
      throw new AgentDocxError(
        "REFERENCE_INVALID",
        `Unsupported asset entry: ${logical}`,
      );
    assets.set(logical, {
      bytes: await readFile(path),
      mediaType: mediaTypeFor(logical),
    });
  }
  return assets;
};

const addDependency = async (
  bytes: Uint8Array | string,
  mediaType: string,
  key: string,
  dependencies: Map<string, { bytes: Uint8Array; mediaType: string }>,
): Promise<RevisionId> => {
  const encoded =
    typeof bytes === "string" ? new TextEncoder().encode(bytes) : bytes;
  dependencies.set(key, { bytes: encoded, mediaType });
  return objectId(encoded);
};

export const snapshotProjectDocument = async (
  opened: OpenedStore,
  document: AgentDocxDocumentConfig,
): Promise<ProjectSnapshot> => {
  const sourcePath = assertWithin(
    opened.projectDirectory,
    document.source,
    "Document source",
  );
  await assertRegularFile(sourcePath, "Document source");
  const sourceBytes = await readFile(sourcePath);
  let source: string;
  try {
    source = new TextDecoder("utf-8", { fatal: true }).decode(sourceBytes);
  } catch {
    throw new AgentDocxError(
      "INPUT_NOT_UTF8",
      `Document source is not UTF-8: ${document.source}`,
    );
  }
  const dependencies = new Map<
    string,
    { bytes: Uint8Array; mediaType: string }
  >();
  const dependencyObjects: Record<string, RevisionId> = {};
  const profile = builtInProfiles[document.profile];
  if (!profile)
    throw new AgentDocxError(
      "PROJECT_INVALID",
      `Unknown profile: ${document.profile}`,
    );
  dependencyObjects.profile = await addDependency(
    canonicalJson(profile),
    "application/json",
    "profile",
    dependencies,
  );
  if (document.rulePack) {
    const pack = builtInRulePacks[document.rulePack];
    if (!pack)
      throw new AgentDocxError(
        "PROJECT_INVALID",
        `Unknown rule pack: ${document.rulePack}`,
      );
    dependencyObjects["rule-pack"] = await addDependency(
      canonicalJson(pack),
      "application/json",
      "rule-pack",
      dependencies,
    );
    const source = await readFile(
      new URL(`../../assets/rules/${pack.sourceExcerpt}`, import.meta.url),
    );
    if (objectId(source) !== pack.sourceSha256)
      throw new AgentDocxError(
        "PROJECT_INVALID",
        `Rule source hash mismatch: ${pack.sourceExcerpt}`,
      );
    dependencyObjects[`rule-source/${pack.sourceExcerpt}`] =
      await addDependency(
        source,
        "text/plain",
        `rule-source/${pack.sourceExcerpt}`,
        dependencies,
      );
  }
  if (document.template) {
    const templatePath = assertWithin(
      opened.projectDirectory,
      document.template,
      "Template",
    );
    await assertRegularFile(templatePath, "Template");
    const bytes = await readFile(templatePath);
    dependencyObjects.template = await addDependency(
      bytes,
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      "template",
      dependencies,
    );
  }
  if (document.assetsDir) {
    const assetsPath = assertWithin(
      opened.projectDirectory,
      document.assetsDir,
      "Assets directory",
    );
    await assertDirectory(assetsPath, "Assets directory");
    for (const [name, asset] of await collectAssets(assetsPath))
      dependencyObjects[`asset/${name}`] = await addDependency(
        asset.bytes,
        asset.mediaType,
        `asset/${name}`,
        dependencies,
      );
  }
  if (document.fontSet) {
    const roles: Readonly<Record<string, string | undefined>> = {
      regular: document.fontSet.regularPath,
      bold: document.fontSet.boldPath,
      italic: document.fontSet.italicPath,
      boldItalic: document.fontSet.boldItalicPath,
    };
    for (const [role, configuredPath] of Object.entries(roles)) {
      if (!configuredPath) continue;
      const fontPath = assertWithin(
        opened.projectDirectory,
        configuredPath,
        `Font ${role}`,
      );
      await assertRegularFile(fontPath, `Font ${role}`);
      dependencyObjects[`font/${role}`] = await addDependency(
        await readFile(fontPath),
        "font/ttf",
        `font/${role}`,
        dependencies,
      );
    }
  }
  const sourceObject = objectId(sourceBytes);
  const documentConfigObject = canonicalObjectId(document);
  const sortedDependencies = Object.fromEntries(
    Object.entries(dependencyObjects).sort(([left], [right]) =>
      left.localeCompare(right),
    ),
  ) as DependencyHashes;
  const workingTreeHash = canonicalObjectId({
    sourceObject,
    documentConfigObject,
    dependencyObjects: sortedDependencies,
  });
  return {
    source,
    sourceObject,
    documentConfigObject,
    dependencyObjects: sortedDependencies,
    dependencyBytes: dependencies,
    workingTreeHash,
  };
};

const bindingPath = (storePath: string): string =>
  resolve(storePath, BINDING_FILE);

const validateBinding = async (opened: OpenedStore): Promise<void> => {
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

const recoverInitialization = async (
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
  } catch {}
  if (committed) {
    await rm(intentPath, { force: true });
    return;
  }
  const binding = await readJsonFile<ProjectBinding>(
    bindingPath(storePath),
  ).catch(() => null);
  if (binding?.projectId === intent.projectId)
    await rm(storePath, { recursive: true, force: true });
  await rm(intentPath, { force: true });
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
  const intent = objectRecord(value, "Export intent");
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
    (attachmentPath !== null && attachmentStagePath === null)
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
    attachmentStagePath,
    docxSha256: digest(intent.docxSha256, "docxSha256"),
    attachmentManifestSha256:
      intent.attachmentManifestSha256 === null
        ? null
        : digest(intent.attachmentManifestSha256, "attachmentManifestSha256"),
    artifactProvenanceSha256,
    artifactStorePath,
    attachmentStorePath,
  };
};

const pathExists = async (path: string): Promise<boolean> => {
  try {
    await lstat(path);
    return true;
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === "ENOENT") return false;
    throw error;
  }
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
      left.name.localeCompare(right.name),
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

const recoverExport = async (
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
  await verifyArtifactDirectory(
    intent.artifactStorePath,
    intent.artifactProvenanceSha256,
    intent.docxSha256,
    intent.attachmentManifestSha256,
  );
  await assertNoSymlinkComponents(intent.outputPath, "Published DOCX");
  const outputExists = await pathExists(intent.outputPath);
  if (outputExists) {
    if (
      (await hashRegularFile(intent.outputPath, "Published DOCX")) !==
      intent.docxSha256
    )
      throw new AgentDocxError(
        "PROJECT_INVALID",
        "Published DOCX hash does not match export intent",
      );
  } else {
    if (!stageIsOwned || !(await pathExists(intent.docxStagePath)))
      throw new AgentDocxError(
        "PROJECT_INVALID",
        "Prepared export is missing its DOCX stage",
      );
    if (
      (await hashRegularFile(intent.docxStagePath, "Staged DOCX")) !==
      intent.docxSha256
    )
      throw new AgentDocxError(
        "PROJECT_INVALID",
        "Staged DOCX hash does not match export intent",
      );
    await link(intent.docxStagePath, intent.outputPath);
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

export const writeExportIntent = async (
  projectDirectory: string,
  intent: ExportIntent,
): Promise<void> => {
  const path = exportIntentPath(projectDirectory);
  await writeExclusive(path, canonicalJson(intent));
};

export const updateExportIntent = async (
  projectDirectory: string,
  intent: ExportIntent,
): Promise<void> => {
  const path = exportIntentPath(projectDirectory);
  const stage = `${path}.${intent.owner}.stage`;
  await writeExclusive(stage, canonicalJson(intent));
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

export const acquireProjectLock = async (
  projectDirectory: string,
): Promise<() => Promise<void>> => {
  const canonicalDirectory = await realpath(projectDirectory);
  const lockPath = resolve(canonicalDirectory, LOCK_NAME);
  try {
    return await lock(canonicalDirectory, {
      realpath: true,
      lockfilePath: lockPath,
      retries: 0,
      stale: 120000,
      update: 30000,
    });
  } catch (error) {
    throw new AgentDocxError("PROJECT_LOCKED", "Project is locked", {
      cause: error instanceof Error ? error.message : String(error),
    });
  }
};

export const openStore = async (manifestPath: string): Promise<OpenedStore> => {
  const absoluteManifestPath = resolve(manifestPath);
  const projectDirectory = dirname(absoluteManifestPath);
  await assertDirectory(projectDirectory, "Project directory");
  const release = await acquireProjectLock(projectDirectory);
  try {
    await recoverInitialization(projectDirectory, absoluteManifestPath);
    await recoverExport(projectDirectory, absoluteManifestPath);
    await assertRegularFile(absoluteManifestPath, "Project manifest");
    const manifest = validateManifest(
      await readJsonFile<unknown>(absoluteManifestPath),
    );
    await validateManifestPaths(projectDirectory, manifest);
    const opened = {
      manifestPath: absoluteManifestPath,
      projectDirectory: await realpath(projectDirectory),
      storePath: resolve(projectDirectory, STORE_DIR),
      manifest,
    };
    await assertDirectory(opened.storePath, "Project store");
    await validateBinding(opened);
    return opened;
  } catch (error) {
    if ((error as AgentDocxError).code === "INPUT_NOT_FOUND")
      throw new AgentDocxError(
        "PROJECT_NOT_FOUND",
        `Project not found: ${absoluteManifestPath}`,
      );
    throw error;
  } finally {
    await release();
  }
};

export const withLockedStore = async <Value>(
  manifestPath: string,
  operation: (opened: OpenedStore) => Promise<Value>,
): Promise<Value> => {
  const absoluteManifestPath = resolve(manifestPath);
  const projectDirectory = dirname(absoluteManifestPath);
  await assertDirectory(projectDirectory, "Project directory");
  const release = await acquireProjectLock(projectDirectory);
  try {
    await recoverInitialization(projectDirectory, absoluteManifestPath);
    await recoverExport(projectDirectory, absoluteManifestPath);
    await assertRegularFile(absoluteManifestPath, "Project manifest");
    const manifest = validateManifest(
      await readJsonFile<unknown>(absoluteManifestPath),
    );
    await validateManifestPaths(projectDirectory, manifest);
    const opened = {
      manifestPath: absoluteManifestPath,
      projectDirectory: await realpath(projectDirectory),
      storePath: resolve(projectDirectory, STORE_DIR),
      manifest,
    };
    await assertDirectory(opened.storePath, "Project store");
    await validateBinding(opened);
    return await operation(opened);
  } catch (error) {
    if ((error as AgentDocxError).code === "INPUT_NOT_FOUND")
      throw new AgentDocxError(
        "PROJECT_NOT_FOUND",
        `Project not found: ${absoluteManifestPath}`,
      );
    throw error;
  } finally {
    await release();
  }
};

export const initializeStore = async (
  manifestPath: string,
  manifest: AgentDocxManifest,
): Promise<OpenedStore> => {
  const absoluteManifestPath = resolve(manifestPath);
  const projectDirectory = dirname(absoluteManifestPath);
  validateManifest(manifest);
  await validateManifestPaths(projectDirectory, manifest);
  await assertDirectory(projectDirectory, "Project directory");
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
  const release = await acquireProjectLock(projectDirectory);
  try {
    await recoverInitialization(projectDirectory, absoluteManifestPath);
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
    await writeExclusive(intentPath, canonicalJson(intent));
    await mkdir(storePath, { recursive: false, mode: 0o700 });
    await writeExclusive(
      bindingPath(storePath),
      canonicalJson({
        schemaVersion: 1,
        projectId: manifest.projectId,
        manifestBasename: manifestName,
      } satisfies ProjectBinding),
    );
    await writeExclusive(absoluteManifestPath, canonicalJson(manifest));
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
    await assertRegularFile(absolutePath, name);
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
    await assertRegularFile(sourcePath, "Document source");
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
    await assertDirectory(absoluteAssets, "Assets directory");
    assetsDir = assertRelativeManifestPath(
      relativePath(projectDirectory, absoluteAssets),
      "Assets directory",
    );
  }
  return {
    id: input.documentId,
    source,
    profile: input.profile,
    ...(input.filingKind ? { filingKind: input.filingKind } : {}),
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
  await writeExclusive(sourcePath, "");
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
    if (objectId(await readFile(path)) !== expectedOld)
      throw new AgentDocxError(
        "WORKING_COPY_CONFLICT",
        `Owned file changed: ${path}`,
      );
  }
  const stage = `${path}.${randomUUID()}.stage`;
  const backup = `${path}.${randomUUID()}.backup`;
  await writeExclusive(stage, bytes);
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
        currentId = objectId(await readFile(path));
      } catch (error) {
        if ((error as NodeJS.ErrnoException).code === "ENOENT") return;
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
      currentId = objectId(await readFile(path));
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code !== "ENOENT") throw error;
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
    if (objectId(await readFile(path)) !== expectedOld)
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
      if (objectId(await readFile(backup)) !== expectedOld)
        throw new AgentDocxError(
          "PROJECT_INVALID",
          `Owned backup mismatch: ${path}`,
        );
      await link(stage, path);
      replacementLinked = true;
    }
    await assertRegularFile(path, "Owned replacement");
    if (objectId(await readFile(path)) !== nextId)
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
  if (objectId(await readFile(path)) !== expected)
    throw new AgentDocxError(
      "WORKING_COPY_CONFLICT",
      `Owned file changed: ${path}`,
    );
  const backup = `${path}.${randomUUID()}.backup`;
  await rename(path, backup);
  try {
    await assertRegularFile(backup, "Owned backup");
    if (objectId(await readFile(backup)) !== expected)
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
      if (objectId(await readFile(path)) !== expected)
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

export const storeSnapshot = async (
  opened: OpenedStore,
  snapshot: ProjectSnapshot,
): Promise<void> => {
  await writeObject(opened.storePath, snapshot.source);
  await writeObject(opened.storePath, canonicalJson(opened.manifest));
  for (const { bytes } of snapshot.dependencyBytes.values())
    await writeObject(opened.storePath, bytes);
};

export const updateManifest = async (
  opened: OpenedStore,
  manifest: AgentDocxManifest,
): Promise<OpenedStore> => {
  validateManifest(manifest);
  await validateManifestPaths(opened.projectDirectory, manifest);
  const stage = `${opened.manifestPath}.${randomUUID()}.stage`;
  await writeExclusive(stage, canonicalJson(manifest));
  try {
    await rename(stage, opened.manifestPath);
  } finally {
    await rm(stage, { force: true });
  }
  return { ...opened, manifest };
};
