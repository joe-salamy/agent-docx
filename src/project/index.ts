import { randomUUID as systemRandomUuid } from "node:crypto";
import { lstat, mkdir, open, rm } from "node:fs/promises";
import { dirname, extname, isAbsolute, relative, resolve, sep } from "node:path";
import { measureNormalizedDocument } from "../renderers/index.js";
import { validateLegalDocument, type ValidationResult } from "../legal/rules.js";
import {
  insertMissingBlockMarkers,
  parseLegalMarkdown,
  type LegalAssetInput,
} from "../legal/parse.js";
import { lowerLegalDocument } from "../legal/lower.js";
import type {
  Actor,
  LegalDocument,
  RevisionId,
  ReviewAnnotation,
} from "../legal/model.js";
import { AgentDocxError, type MeasurementResult } from "../types.js";
import {
  createChangeSet,
  createRevisionDelta,
  defaultAttribution,
  rebaseOpenAnnotations,
  visibleTextForBlock,
} from "../revisions/diff.js";
import type {
  Change,
  ChangeSet,
  ResolutionRecord,
  RevisionMutationResult,
  RevisionPage,
  RevisionRecord,
} from "../revisions/types.js";
import type {
  AddReviewInput,
  AgentDocxDocumentConfig,
  AgentDocxManifest,
  AgentDocxProject,
  ConfigureDocumentInput,
  DocumentSnapshot,
  ProjectDocumentInput,
  ProjectMeasureOptions,
  ProjectMeasurementResult,
  ProjectState,
  ResolveChangesInput,
  ResolveReviewInput,
} from "./contracts.js";
import {
  clearExportIntent,
  completeExportIntent,
  createEmptySource,
  canonicalObjectId,
  documentConfigFromInput,
  initializeStore,
  canonicalJson,
  objectId,
  openStore,
  readHead,
  readObject,
  readRevisionJson,
  replaceOwnedFile,
  snapshotProjectDocument,
  storeSnapshot,
  updateExportIntent,
  updateManifest,
  withLockedStore,
  writeHead,
  writeObject,
  writeRevisionJson,
  type ExportIntent,
  type OpenedStore,
  type ProjectSnapshot,
} from "./store.js";
import type {
  CompiledDocx,
  ExportDocxInput,
  GeneratedAttachmentBundle,
  ImportDocxInput,
  DocxImportResult,
} from "../docx/contracts.js";
import type { DraftGuidance, PatchEvaluation, SourcePatch } from "../draft/types.js";
import {
  compileMarkdown,
  createSemanticManifest,
  semanticDocumentProjection,
} from "../docx/compile.js";
import { inspectDocxMaterial } from "../docx/import.js";
import { generateRedlineDocx } from "../docx/redline.js";

export type ProjectRuntimeOptions = {
  clock?: () => Date;
  randomUUID?: () => string;
};

type RevisionMaterial = {
  revision: RevisionRecord;
  source: string;
  config: AgentDocxDocumentConfig;
  document: LegalDocument;
  annotations: readonly ReviewAnnotation[];
};

const version = "0.1.0";

const sourcePathFor = (opened: OpenedStore, config: AgentDocxDocumentConfig): string =>
  resolve(opened.projectDirectory, config.source);

const documentById = (
  manifest: AgentDocxManifest,
  documentId: string,
): AgentDocxDocumentConfig => {
  const document = manifest.documents.find((entry) => entry.id === documentId);
  if (!document)
    throw new AgentDocxError("DOCUMENT_NOT_FOUND", `Document not found: ${documentId}`);
  return document;
};


const sourceAssets = (
  snapshot: ProjectSnapshot,
): Readonly<Record<string, LegalAssetInput>> => {
  const assets: Record<string, LegalAssetInput> = {};
  for (const [key, value] of snapshot.dependencyBytes) {
    if (!key.startsWith("asset/")) continue;
    assets[key.slice("asset/".length)] = value;
  }
  return assets;
};

const importedAssetDestinations = (
  opened: OpenedStore,
  config: AgentDocxDocumentConfig,
  assets: Readonly<Record<string, LegalAssetInput>>,
): readonly { source: string; path: string; bytes: Uint8Array }[] => {
  const entries = Object.entries(assets).sort(([left], [right]) =>
    left.localeCompare(right),
  );
  if (entries.length === 0) return [];
  if (!config.assetsDir)
    throw new AgentDocxError(
      "DOCX_IMPORT_UNSUPPORTED",
      "Imported DOCX assets require the target document to configure assetsDir",
    );
  const root = resolve(opened.projectDirectory, config.assetsDir);
  return entries.map(([source, asset]) => {
    if (
      source.length === 0 ||
      source.startsWith("/") ||
      source.includes("\\") ||
      source.split("/").some((part) => part === "" || part === "." || part === "..")
    )
      throw new AgentDocxError(
        "DOCX_IMPORT_UNSUPPORTED",
        `Imported asset has an unsafe source path: ${source}`,
      );
    const path = resolve(root, source);
    const contained = relative(root, path);
    if (
      contained === "" ||
      isAbsolute(contained) ||
      contained === ".." ||
      contained.startsWith(`..${sep}`)
    )
      throw new AgentDocxError(
        "DOCX_IMPORT_UNSUPPORTED",
        `Imported asset escapes assetsDir: ${source}`,
      );
    return { source, path, bytes: asset.bytes };
  });
};

const materializeSourceMarkers = async (
  opened: OpenedStore,
  config: AgentDocxDocumentConfig,
): Promise<void> => {
  const snapshot = await snapshotProjectDocument(opened, config);
  const marked = insertMissingBlockMarkers(snapshot.source, {
    projectId: opened.manifest.projectId,
    documentId: config.id,
    metadata: config.metadata,
    chrome: config.chrome,
    assets: sourceAssets(snapshot),
  });
  if (marked !== snapshot.source)
    await replaceOwnedFile(
      sourcePathFor(opened, config),
      snapshot.sourceObject,
      marked,
    );
};

const sourceFontSet = (
  config: AgentDocxDocumentConfig,
  snapshot: ProjectSnapshot,
) => {
  if (!config.fontSet) return undefined;
  const regular = snapshot.dependencyBytes.get("font/regular");
  if (!regular)
    throw new AgentDocxError("PROJECT_INVALID", "Configured regular font is missing");
  const bold = snapshot.dependencyBytes.get("font/bold");
  const italic = snapshot.dependencyBytes.get("font/italic");
  const boldItalic = snapshot.dependencyBytes.get("font/boldItalic");
  return {
    family: config.fontSet.family,
    regular: regular.bytes,
    ...(bold ? { bold: bold.bytes } : {}),
    ...(italic ? { italic: italic.bytes } : {}),
    ...(boldItalic ? { boldItalic: boldItalic.bytes } : {}),
  };
};

const storedMediaType = (key: string): string => {
  if (key === "template")
    return "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
  if (key.startsWith("font/")) return "font/ttf";
  if (key.endsWith(".png")) return "image/png";
  if (key.endsWith(".jpg") || key.endsWith(".jpeg")) return "image/jpeg";
  if (key.endsWith(".pdf")) return "application/pdf";
  return key.startsWith("rule-source/") ? "text/plain" : "application/octet-stream";
};

const serializableMeasurement = (
  measurement: MeasurementResult,
): Omit<MeasurementResult, "generatedDocx"> => {
  const { generatedDocx: _generatedDocx, ...result } = measurement;
  return result;
};

const attachmentDirectoryFor = (output: string): string => {
  const path = resolve(output);
  const extension = extname(path);
  return `${extension.toLowerCase() === ".docx" ? path.slice(0, -extension.length) : path}.attachments`;
};

const writeExclusiveFile = async (
  path: string,
  bytes: Uint8Array | string,
): Promise<void> => {
  const handle = await open(path, "wx", 0o600);
  try {
    await handle.writeFile(bytes);
  } finally {
    await handle.close();
  }
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

const pathsOverlap = (left: string, right: string): boolean =>
  left === right || left.startsWith(`${right}${sep}`) || right.startsWith(`${left}${sep}`);

const objectStorePath = (storePath: string, object: string): string =>
  resolve(
    storePath,
    "objects",
    "sha256",
    object.slice("sha256:".length, "sha256:".length + 2),
    object.slice("sha256:".length + 2),
  );

const exportOwnedPaths = (
  opened: OpenedStore,
  config: AgentDocxDocumentConfig,
): readonly string[] =>
  [
    opened.manifestPath,
    opened.storePath,
    sourcePathFor(opened, config),
    config.template ? resolve(opened.projectDirectory, config.template) : null,
    config.assetsDir ? resolve(opened.projectDirectory, config.assetsDir) : null,
    config.fontSet?.regularPath
      ? resolve(opened.projectDirectory, config.fontSet.regularPath)
      : null,
    config.fontSet?.boldPath
      ? resolve(opened.projectDirectory, config.fontSet.boldPath)
      : null,
    config.fontSet?.italicPath
      ? resolve(opened.projectDirectory, config.fontSet.italicPath)
      : null,
    config.fontSet?.boldItalicPath
      ? resolve(opened.projectDirectory, config.fontSet.boldItalicPath)
      : null,
  ].filter((path): path is string => path !== null);

const assertExportDestination = async (
  opened: OpenedStore,
  config: AgentDocxDocumentConfig,
  output: string,
): Promise<{ output: string; attachment: string }> => {
  const absoluteOutput = resolve(output);
  const relativeOutput = relative(opened.projectDirectory, absoluteOutput);
  if (
    relativeOutput === ".." ||
    relativeOutput.startsWith(`..${sep}`) ||
    isAbsolute(relativeOutput)
  )
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      "DOCX output must be inside the project directory",
    );
  const attachment = attachmentDirectoryFor(absoluteOutput);
  const ownedPaths = exportOwnedPaths(opened, config);
  if (ownedPaths.some((owned) => pathsOverlap(absoluteOutput, owned))) {
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      "DOCX output overlaps a project source, dependency, or control path",
    );
  }
  await assertRegularDirectory(dirname(absoluteOutput), "DOCX output parent");
  if (await pathExists(absoluteOutput))
    throw new AgentDocxError("OUTPUT_EXISTS", `DOCX output already exists: ${absoluteOutput}`);
  if (await pathExists(attachment))
    throw new AgentDocxError(
      "OUTPUT_EXISTS",
      `Attachment bundle already exists: ${attachment}`,
    );
  return { output: absoluteOutput, attachment };
};

const assertRegularDirectory = async (path: string, label: string): Promise<void> => {
  let entry;
  try {
    entry = await lstat(path);
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === "ENOENT")
      throw new AgentDocxError("INPUT_NOT_FOUND", `${label} does not exist: ${path}`);
    throw error;
  }
  if (!entry.isDirectory() || entry.isSymbolicLink())
    throw new AgentDocxError("INVALID_ARGUMENT", `${label} is not a regular directory: ${path}`);
};

const createExportStage = async (
  stagePath: string,
  owner: string,
  projectId: string,
  manifestPath: string,
): Promise<void> => {
  await mkdir(stagePath, { mode: 0o700 });
  await writeExclusiveFile(
    resolve(stagePath, "owner.json"),
    canonicalJson({ schemaVersion: 1, owner, projectId, manifestPath }),
  );
};

const writeAttachmentStage = async (
  stagePath: string,
  bundle: GeneratedAttachmentBundle,
): Promise<void> => {
  const attachmentRoot = resolve(stagePath, "attachments");
  await mkdir(attachmentRoot, { mode: 0o700 });
  for (const entry of bundle.manifest.entries) {
    if (!entry.payloadPath.startsWith("files/") || entry.payloadPath.includes("\\"))
      throw new AgentDocxError("PROJECT_INVALID", `Invalid attachment payload path: ${entry.payloadPath}`);
    const payload = bundle.files[entry.name];
    if (!payload)
      throw new AgentDocxError("PROJECT_INVALID", `Missing attachment bytes: ${entry.name}`);
    const destination = resolve(attachmentRoot, entry.payloadPath);
    const relativePayload = relative(attachmentRoot, destination);
    if (
      relativePayload === ".." ||
      relativePayload.startsWith(`..${sep}`) ||
      isAbsolute(relativePayload)
    )
      throw new AgentDocxError("PROJECT_INVALID", "Attachment payload escapes its bundle");
    await mkdir(dirname(destination), { recursive: true, mode: 0o700 });
    await writeExclusiveFile(destination, payload.bytes);
  }
  await writeExclusiveFile(
    resolve(attachmentRoot, "manifest.json"),
    canonicalJson(bundle.manifest),
  );
};


const documentFor = (
  source: string,
  config: AgentDocxDocumentConfig,
  snapshot: ProjectSnapshot,
  projectId: string,
  annotations: readonly ReviewAnnotation[],
  requireMarkers: boolean,
): LegalDocument =>
  parseLegalMarkdown(source, {
    projectId,
    documentId: config.id,
    metadata: config.metadata,
    chrome: config.chrome,
    assets: sourceAssets(snapshot),
    annotations,
    requireMarkers,
  }).document;

const mutationError = (error: unknown) => {
  if (error instanceof AgentDocxError)
    return { code: error.code, message: error.message, ...(error.details ? { details: error.details } : {}) };
  return {
    code: "PATCH_INVALID" as const,
    message: error instanceof Error ? error.message : String(error),
  };
};

type RawReplacement = {
  start: number;
  end: number;
  expectedText: string;
  replacement: string;
};

const rejectedSourceReplacements = (
  source: string,
  changes: readonly Change[],
  decisions: Readonly<Record<`c_${string}`, "accept" | "reject">>,
): RawReplacement[] => {
  const replacements: RawReplacement[] = [];
  for (const change of changes) {
    if (decisions[change.id] !== "reject") continue;
    if (change.kind === "insert-block" || change.kind === "insert-text") {
      replacements.push({
        start: change.newSource.start,
        end: change.newSource.end,
        expectedText: change.newSource.text,
        replacement: "",
      });
      continue;
    }
    if (change.kind === "replace-text") {
      replacements.push({
        start: change.newSource.start,
        end: change.newSource.end,
        expectedText: change.newSource.text,
        replacement: change.oldSource.text,
      });
      continue;
    }
    if (change.kind === "replace-block") {
      replacements.push({
        start: change.newBlock.position.start.offset,
        end: change.newBlock.position.end.offset,
        expectedText: change.newBlock.sourceText,
        replacement: change.oldBlock.sourceText,
      });
      continue;
    }
    throw new AgentDocxError(
      "CHANGESET_INVALID",
      `Cannot safely reject ${change.kind} without an exact head source range`,
    );
  }
  for (const replacement of replacements)
    if (source.slice(replacement.start, replacement.end) !== replacement.expectedText)
      throw new AgentDocxError(
        "REVISION_CONFLICT",
        "Change-set head source no longer matches its recorded range",
      );
  const ordered = [...replacements].sort((left, right) => right.start - left.start);
  for (const [index, replacement] of ordered.entries()) {
    const next = ordered[index + 1];
    if (next && replacement.start < next.end)
      throw new AgentDocxError("CHANGESET_INVALID", "Rejected source changes overlap");
  }
  return ordered;
};

const isUtf16Boundary = (text: string, offset: number): boolean =>
  Number.isInteger(offset) &&
  offset >= 0 &&
  offset <= text.length &&
  (offset === 0 ||
    offset === text.length ||
    !(
      text.charCodeAt(offset - 1) >= 0xd800 &&
      text.charCodeAt(offset - 1) <= 0xdbff &&
      text.charCodeAt(offset) >= 0xdc00 &&
      text.charCodeAt(offset) <= 0xdfff
    ));

class Project implements AgentDocxProject {
  constructor(
    private readonly manifestPath: string,
    private readonly clock: () => Date,
    private readonly randomUuid: () => string,
  ) {}

  private async currentRevision(
    opened: OpenedStore,
    documentId: string,
    selector: RevisionId | "HEAD",
  ): Promise<RevisionRecord> {
    const head = await readHead(opened.storePath, documentId);
    const requested = selector === "HEAD" ? head : selector;
    if (!requested)
      throw new AgentDocxError("REVISION_NOT_FOUND", `Document has no revision: ${documentId}`);
    const visited = new Set<string>();
    const pending: RevisionId[] = head ? [head] : [];
    while (pending.length > 0) {
      const next = pending.pop()!;
      if (visited.has(next)) continue;
      visited.add(next);
      const record = await readRevisionJson<RevisionRecord>(opened.storePath, next);
      if (record.documentId !== documentId)
        throw new AgentDocxError("PROJECT_INVALID", `Revision belongs to another document: ${next}`);
      if (record.id === requested) return record;
      pending.push(...record.parents);
    }
    throw new AgentDocxError("REVISION_NOT_FOUND", `Revision not found: ${requested}`);
  }

  private async isFirstParentAncestor(
    opened: OpenedStore,
    ancestor: RevisionId,
    descendant: RevisionRecord,
  ): Promise<boolean> {
    let current: RevisionRecord | null = descendant;
    while (current) {
      if (current.id === ancestor) return true;
      const parent: RevisionId | undefined = current.parents[0];
      current = parent
        ? await readRevisionJson<RevisionRecord>(opened.storePath, parent)
        : null;
    }
    return false;
  }

  private async materialFor(
    opened: OpenedStore,
    record: RevisionRecord,
  ): Promise<RevisionMaterial> {
    const source = new TextDecoder("utf-8", { fatal: true }).decode(
      await readObject(opened.storePath, record.sourceObject),
    );
    const config = JSON.parse(
      new TextDecoder("utf-8", { fatal: true }).decode(
        await readObject(opened.storePath, record.documentConfigObject),
      ),
    ) as AgentDocxDocumentConfig;
    const document = JSON.parse(
      new TextDecoder("utf-8", { fatal: true }).decode(
        await readObject(opened.storePath, record.legalDocumentObject),
      ),
    ) as LegalDocument;
    const annotations = JSON.parse(
      new TextDecoder("utf-8", { fatal: true }).decode(
        await readObject(opened.storePath, record.annotationsObject),
      ),
    ) as ReviewAnnotation[];
    return { revision: record, source, config, document, annotations };
  }

  private async snapshotForMaterial(
    opened: OpenedStore,
    material: RevisionMaterial,
  ): Promise<ProjectSnapshot> {
    const dependencyBytes = new Map<
      string,
      { bytes: Uint8Array; mediaType: string }
    >();
    for (const [key, id] of Object.entries(material.revision.dependencyObjects))
      dependencyBytes.set(key, {
        bytes: await readObject(opened.storePath, id),
        mediaType: storedMediaType(key),
      });
    return {
      source: material.source,
      sourceObject: material.revision.sourceObject,
      documentConfigObject: material.revision.documentConfigObject,
      dependencyObjects: material.revision.dependencyObjects,
      dependencyBytes,
      workingTreeHash: material.revision.workingTreeHash,
    };
  }

  private async annotationsForHead(
    opened: OpenedStore,
    documentId: string,
  ): Promise<readonly ReviewAnnotation[]> {
    const head = await readHead(opened.storePath, documentId);
    if (!head) return [];
    const record = await readRevisionJson<RevisionRecord>(opened.storePath, head);
    return JSON.parse(
      new TextDecoder("utf-8", { fatal: true }).decode(
        await readObject(opened.storePath, record.annotationsObject),
      ),
    ) as ReviewAnnotation[];
  }

  private async measureSnapshot(
    config: AgentDocxDocumentConfig,
    snapshot: ProjectSnapshot,
    document: LegalDocument,
    revision: RevisionId | null,
    options: ProjectMeasureOptions = {},
  ): Promise<ProjectMeasurementResult> {
    const measurement = await measureNormalizedDocument(lowerLegalDocument(document), {
      ...options,
      profile: config.profile,
      filingKind: config.filingKind,
      chrome: config.chrome,
      fontSet: sourceFontSet(config, snapshot),
    });
    return {
      ...measurement,
      documentId: config.id,
      revision,
      workingTreeHash: snapshot.workingTreeHash,
      dependencyObjects: snapshot.dependencyObjects,
    };
  }

  private async prepareWorkingDocument(
    opened: OpenedStore,
    config: AgentDocxDocumentConfig,
    snapshot: ProjectSnapshot,
    annotations: readonly ReviewAnnotation[],
    materializeMarkers: boolean,
  ): Promise<{ source: string; snapshot: ProjectSnapshot; document: LegalDocument }> {
    let source = snapshot.source;
    if (materializeMarkers) {
      const marked = insertMissingBlockMarkers(source, {
        projectId: opened.manifest.projectId,
        documentId: config.id,
        metadata: config.metadata,
        chrome: config.chrome,
        assets: sourceAssets(snapshot),
      });
      if (marked !== source) {
        await replaceOwnedFile(
          sourcePathFor(opened, config),
          objectId(source),
          marked,
        );
        snapshot = await snapshotProjectDocument(opened, config);
        source = snapshot.source;
      }
    }
    return {
      source,
      snapshot,
      document: documentFor(
        source,
        config,
        snapshot,
        opened.manifest.projectId,
        annotations,
        materializeMarkers,
      ),
    };
  }

  private async commitLocked(
    opened: OpenedStore,
    config: AgentDocxDocumentConfig,
    snapshot: ProjectSnapshot,
    document: LegalDocument,
    annotations: readonly ReviewAnnotation[],
    base: RevisionId | null,
    author: Actor,
    message: string,
    resolution?: ResolutionRecord,
    preserveAnnotations = false,
  ): Promise<RevisionMutationResult> {
    const head = await readHead(opened.storePath, config.id);
    if (base === null ? head !== null : head !== base)
      throw new AgentDocxError("REVISION_CONFLICT", "Revision base does not match current head");
    const parent = head ? await readRevisionJson<RevisionRecord>(opened.storePath, head) : null;
    const parentMaterial = parent
      ? await this.materialFor(opened, parent)
      : null;
    const committedAnnotations =
      parentMaterial && !preserveAnnotations
        ? rebaseOpenAnnotations(parentMaterial.document, document, annotations)
        : annotations;
    const committedDocument = {
      ...document,
      annotations: committedAnnotations,
    };
    await storeSnapshot(opened, snapshot);
    const sourceObject = await writeObject(opened.storePath, snapshot.source);
    const documentConfigObject = await writeObject(
      opened.storePath,
      JSON.stringify(config),
    );
    const legalDocumentObject = await writeObject(
      opened.storePath,
      JSON.stringify(committedDocument),
    );
    const annotationsObject = await writeObject(
      opened.storePath,
      JSON.stringify(committedAnnotations),
    );
    const createdAt = this.clock().toISOString();
    const deltaObject = parent
      ? await writeObject(
          opened.storePath,
          canonicalJson(
            createRevisionDelta(
              parent,
              config.id,
              parentMaterial!.document,
              committedDocument,
              parentMaterial!.annotations,
              committedAnnotations,
              defaultAttribution(author, createdAt),
            ),
          ),
        )
      : undefined;
    const resolutionObject = resolution
      ? await writeObject(opened.storePath, canonicalJson(resolution))
      : undefined;
    const record = {
      schemaVersion: 1 as const,
      documentId: config.id,
      parents: parent ? [parent.id] : [],
      createdAt,
      author,
      message,
      sourceObject,
      documentConfigObject,
      dependencyObjects: snapshot.dependencyObjects,
      workingTreeHash: snapshot.workingTreeHash,
      legalDocumentObject,
      annotationsObject,
      ...(deltaObject ? { deltaObject } : {}),
      ...(resolutionObject ? { resolutionObject } : {}),
      tool: { name: "agent-docx" as const, version, schemaVersion: 1 as const },
    };
    const revisionId = await writeRevisionJson(opened.storePath, record);
    const revision = { ...record, id: revisionId } satisfies RevisionRecord;
    const measurement = await this.measureSnapshot(
      config,
      snapshot,
      committedDocument,
      revisionId,
    );
    const validation = validateLegalDocument(committedDocument, {
      revision: revisionId,
      rulePack: config.rulePack,
      filingKind: config.filingKind,
      measurement: serializableMeasurement(measurement),
    });
    await writeHead(opened.storePath, config.id, revisionId);
    return {
      schemaVersion: 1,
      revision,
      head: revisionId,
      sourceSha256: sourceObject,
      workingTreeHash: snapshot.workingTreeHash,
      measurement,
      validation,
    };
  }

  async getState(): Promise<ProjectState> {
    const opened = await openStore(this.manifestPath);
    const documents = await Promise.all(
      opened.manifest.documents.map(async (config) => {
        const snapshot = await snapshotProjectDocument(opened, config);
        const head = await readHead(opened.storePath, config.id);
        const record = head
          ? await readRevisionJson<RevisionRecord>(opened.storePath, head)
          : null;
        return {
          documentId: config.id,
          head,
          headWorkingTreeHash: record?.workingTreeHash ?? null,
          workingTreeHash: snapshot.workingTreeHash,
          sourceSha256: snapshot.sourceObject,
          documentConfigSha256: snapshot.documentConfigObject,
          dependencyObjects: snapshot.dependencyObjects,
          matchesHead: {
            source: record?.sourceObject === snapshot.sourceObject,
            documentConfig: record?.documentConfigObject === snapshot.documentConfigObject,
            dependencies:
              record !== null &&
              JSON.stringify(record.dependencyObjects) ===
                JSON.stringify(snapshot.dependencyObjects),
            all: record?.workingTreeHash === snapshot.workingTreeHash,
          },
        };
      }),
    );
    return {
      schemaVersion: 1,
      manifestPath: this.manifestPath,
      manifest: opened.manifest,
      documents,
    };
  }

  async addDocument(
    input: ProjectDocumentInput & { makeDefault?: boolean },
  ): Promise<ProjectState> {
    await withLockedStore(this.manifestPath, async (opened) => {
      if (opened.manifest.documents.some((document) => document.id === input.documentId))
        throw new AgentDocxError("DOCUMENT_EXISTS", `Document already exists: ${input.documentId}`);
      const config = await documentConfigFromInput(opened.projectDirectory, input);
      if (input.createSource) await createEmptySource(opened.projectDirectory, input.source);
      const manifest: AgentDocxManifest = {
        ...opened.manifest,
        defaultDocument: input.makeDefault ? config.id : opened.manifest.defaultDocument,
        documents: [...opened.manifest.documents, config],
      };
      await updateManifest(opened, manifest);
      await materializeSourceMarkers(opened, config);
    });
    return this.getState();
  }

  async configureDocument(
    documentId: string,
    input: ConfigureDocumentInput,
  ): Promise<RevisionMutationResult> {
    return withLockedStore(this.manifestPath, async (opened) => {
      const current = documentById(opened.manifest, documentId);
      const head = await readHead(opened.storePath, documentId);
      const base = input.baseRevision === "HEAD" ? head : input.baseRevision;
      if (base === null ? head !== null : head !== base)
        throw new AgentDocxError("REVISION_CONFLICT", "Document configuration base does not match head");
      if (Object.keys(input.changes).length === 0)
        throw new AgentDocxError("INVALID_ARGUMENT", "Document configuration changes are required");
      const next: AgentDocxDocumentConfig = { ...current };
      if (input.changes.profile !== undefined) next.profile = input.changes.profile;
      if (input.changes.metadata !== undefined) next.metadata = input.changes.metadata;
      if (input.changes.filingKind === null) delete next.filingKind;
      else if (input.changes.filingKind !== undefined)
        next.filingKind = input.changes.filingKind;
      if (input.changes.rulePack === null) delete next.rulePack;
      else if (input.changes.rulePack !== undefined)
        next.rulePack = input.changes.rulePack;
      if (input.changes.template === null) delete next.template;
      else if (input.changes.template !== undefined)
        next.template = input.changes.template;
      if (input.changes.assetsDir === null) delete next.assetsDir;
      else if (input.changes.assetsDir !== undefined)
        next.assetsDir = input.changes.assetsDir;
      if (input.changes.fontSet === null) delete next.fontSet;
      else if (input.changes.fontSet !== undefined)
        next.fontSet = input.changes.fontSet;
      if (input.changes.chrome === null) delete next.chrome;
      else if (input.changes.chrome !== undefined)
        next.chrome = input.changes.chrome;
      const manifest: AgentDocxManifest = {
        ...opened.manifest,
        documents: opened.manifest.documents.map((document) =>
          document.id === documentId ? next : document,
        ),
      };
      const updated = await updateManifest(opened, manifest);
      const snapshot = await snapshotProjectDocument(updated, next);
      const annotations = await this.annotationsForHead(updated, documentId);
      const prepared = await this.prepareWorkingDocument(
        updated,
        next,
        snapshot,
        annotations,
        true,
      );
      return this.commitLocked(
        updated,
        next,
        prepared.snapshot,
        prepared.document,
        annotations,
        base,
        input.author,
        input.message,
      );
    });
  }

  async getDocument(
    documentId: string,
    revision?: RevisionId | "HEAD",
  ): Promise<DocumentSnapshot> {
    const opened = await openStore(this.manifestPath);
    const config = documentById(opened.manifest, documentId);
    const head = await readHead(opened.storePath, documentId);
    if (revision !== undefined) {
      const record = await this.currentRevision(opened, documentId, revision);
      const material = await this.materialFor(opened, record);
      return {
        schemaVersion: 1,
        documentId,
        revision: record.id,
        head,
        source: material.source,
        sourceSha256: record.sourceObject,
        workingTreeHash: record.workingTreeHash,
        documentConfig: material.config,
        dependencyObjects: record.dependencyObjects,
        document: material.document,
        annotations: material.annotations,
      };
    }
    const snapshot = await snapshotProjectDocument(opened, config);
    const annotations = await this.annotationsForHead(opened, documentId);
    const document = documentFor(
      snapshot.source,
      config,
      snapshot,
      opened.manifest.projectId,
      annotations,
      false,
    );
    return {
      schemaVersion: 1,
      documentId,
      revision: null,
      head,
      source: snapshot.source,
      sourceSha256: snapshot.sourceObject,
      workingTreeHash: snapshot.workingTreeHash,
      documentConfig: config,
      dependencyObjects: snapshot.dependencyObjects,
      document,
      annotations,
    };
  }

  async checkpoint(
    documentId: string,
    input: { baseRevision: RevisionId | "HEAD" | null; author: Actor; message: string },
  ): Promise<RevisionMutationResult> {
    return withLockedStore(this.manifestPath, async (opened) => {
      const config = documentById(opened.manifest, documentId);
      const head = await readHead(opened.storePath, documentId);
      const base = input.baseRevision === "HEAD" ? head : input.baseRevision;
      const snapshot = await snapshotProjectDocument(opened, config);
      const annotations = await this.annotationsForHead(opened, documentId);
      const prepared = await this.prepareWorkingDocument(
        opened,
        config,
        snapshot,
        annotations,
        true,
      );
      return this.commitLocked(
        opened,
        config,
        prepared.snapshot,
        prepared.document,
        annotations,
        base,
        input.author,
        input.message,
      );
    });
  }

  async listRevisions(
    documentId: string,
    input: { limit?: number; cursor?: RevisionId } = {},
  ): Promise<RevisionPage> {
    const opened = await openStore(this.manifestPath);
    const head = await readHead(opened.storePath, documentId);
    const reachable = new Map<RevisionId, RevisionRecord>();
    const pending = head ? [head] : [];
    while (pending.length > 0) {
      const current = pending.pop()!;
      if (reachable.has(current)) continue;
      const record = await readRevisionJson<RevisionRecord>(opened.storePath, current);
      if (record.documentId !== documentId)
        throw new AgentDocxError("PROJECT_INVALID", "Revision graph crosses documents");
      reachable.set(current, record);
      pending.push(...record.parents);
    }
    const ordered = [...reachable.values()].sort(
      (left, right) =>
        right.createdAt.localeCompare(left.createdAt) || left.id.localeCompare(right.id),
    );
    const limit = input.limit ?? 100;
    if (!Number.isInteger(limit) || limit < 1 || limit > 1000)
      throw new AgentDocxError("INVALID_ARGUMENT", "Revision limit must be 1 through 1000");
    const start = input.cursor
      ? Math.max(0, ordered.findIndex((record) => record.id === input.cursor) + 1)
      : 0;
    if (input.cursor && start === 0)
      throw new AgentDocxError("REVISION_NOT_FOUND", `Revision cursor not found: ${input.cursor}`);
    const items = ordered.slice(start, start + limit);
    return {
      schemaVersion: 1,
      items,
      nextCursor: ordered[start + limit]?.id ?? null,
    };
  }

  async getRevision(documentId: string, revision: RevisionId | "HEAD"): Promise<RevisionRecord> {
    const opened = await openStore(this.manifestPath);
    return this.currentRevision(opened, documentId, revision);
  }

  async diff(
    documentId: string,
    base: RevisionId | "HEAD",
    head: RevisionId | "HEAD",
  ): Promise<ChangeSet> {
    const opened = await openStore(this.manifestPath);
    const baseRecord = await this.currentRevision(opened, documentId, base);
    const headRecord = await this.currentRevision(opened, documentId, head);
    const baseMaterial = await this.materialFor(opened, baseRecord);
    const headMaterial = await this.materialFor(opened, headRecord);
    return createChangeSet(
      documentId,
      baseRecord.id,
      headRecord.id,
      baseMaterial.document,
      headMaterial.document,
      baseMaterial.annotations,
      headMaterial.annotations,
      defaultAttribution(headRecord.author, headRecord.createdAt),
    );
  }

  async addReview(
    documentId: string,
    input: AddReviewInput,
  ): Promise<RevisionMutationResult> {
    return withLockedStore(this.manifestPath, async (opened) => {
      const record = await this.currentRevision(opened, documentId, input.revision);
      const head = await readHead(opened.storePath, documentId);
      if (head !== record.id)
        throw new AgentDocxError("REVISION_CONFLICT", "Review must target the current head");
      const material = await this.materialFor(opened, record);
      const block = [
        ...material.document.blocks,
        ...material.document.footnotes,
      ].find((entry) => entry.id === input.blockId);
      if (!block)
        throw new AgentDocxError("REFERENCE_INVALID", `Block not found: ${input.blockId}`);
      if (
        input.range &&
        (!isUtf16Boundary(visibleTextForBlock(block), input.range.start) ||
          !isUtf16Boundary(visibleTextForBlock(block), input.range.end) ||
          input.range.start > input.range.end)
      )
        throw new AgentDocxError(
          "ANNOTATION_CONFLICT",
          "Review range must be a code-point-safe range within its block",
        );
      if (input.range && input.range.start >= input.range.end)
        throw new AgentDocxError(
          "ANNOTATION_CONFLICT",
          "Review text range must select at least one code point",
        );
      if (
        input.range &&
        material.annotations.some(
          (annotation) =>
            annotation.status === "open" &&
            annotation.blockId === input.blockId &&
            annotation.range !== undefined &&
            input.range!.start < annotation.range.end &&
            annotation.range.start < input.range!.end,
        )
      )
        throw new AgentDocxError(
          "ANNOTATION_CONFLICT",
          "Open review text ranges must not overlap within a block",
        );
      if (
        input.range &&
        ![
          "paragraph",
          "blockquote",
          "heading",
          "numbered-paragraph",
        ].includes(block.kind) &&
        !(block.kind === "footnote" && block.paragraphs.length === 1)
      )
        throw new AgentDocxError(
          "ANNOTATION_CONFLICT",
          "Review ranges must stay within one source-mapped paragraph",
        );
      const annotationId = `a_${this.randomUuid()}` as `a_${string}`;
      if (!/^a_[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/.test(annotationId))
        throw new AgentDocxError("PROJECT_INVALID", "Runtime randomUUID did not return a UUIDv4");
      const annotation: ReviewAnnotation = {
        id: annotationId,
        blockId: input.blockId,
        ...(input.range ? { range: input.range } : {}),
        author: input.author,
        createdAt: this.clock().toISOString(),
        message: input.message,
        status: "open",
      };
      const config = documentById(opened.manifest, documentId);
      const snapshot = await snapshotProjectDocument(opened, config);
      if (snapshot.workingTreeHash !== record.workingTreeHash)
        throw new AgentDocxError(
          "WORKING_COPY_CONFLICT",
          "Working copy differs from the review revision",
        );
      return this.commitLocked(
        opened,
        config,
        snapshot,
        material.document,
        [...material.annotations, annotation],
        record.id,
        input.author,
        input.message,
      );
    });
  }

  async resolveReview(
    documentId: string,
    input: ResolveReviewInput,
  ): Promise<RevisionMutationResult> {
    return withLockedStore(this.manifestPath, async (opened) => {
      const record = await this.currentRevision(opened, documentId, input.revision);
      const head = await readHead(opened.storePath, documentId);
      if (head !== record.id)
        throw new AgentDocxError("REVISION_CONFLICT", "Review must target the current head");
      const material = await this.materialFor(opened, record);
      const annotations = material.annotations.map((annotation) =>
        annotation.id === input.annotationId
          ? { ...annotation, status: "resolved" as const }
          : annotation,
      );
      if (annotations.every((annotation) => annotation.id !== input.annotationId))
        throw new AgentDocxError("REFERENCE_INVALID", `Annotation not found: ${input.annotationId}`);
      const config = documentById(opened.manifest, documentId);
      const snapshot = await snapshotProjectDocument(opened, config);
      if (snapshot.workingTreeHash !== record.workingTreeHash)
        throw new AgentDocxError(
          "WORKING_COPY_CONFLICT",
          "Working copy differs from the review revision",
        );
      return this.commitLocked(
        opened,
        config,
        snapshot,
        material.document,
        annotations,
        record.id,
        input.author,
        input.message,
      );
    });
  }

  async measure(
    documentId: string,
    revision?: RevisionId | "HEAD",
    options: ProjectMeasureOptions = {},
  ): Promise<ProjectMeasurementResult> {
    const opened = await openStore(this.manifestPath);
    if (revision === undefined) {
      const config = documentById(opened.manifest, documentId);
      const snapshot = await snapshotProjectDocument(opened, config);
      const annotations = await this.annotationsForHead(opened, documentId);
      const document = documentFor(
        snapshot.source,
        config,
        snapshot,
        opened.manifest.projectId,
        annotations,
        false,
      );
      return this.measureSnapshot(config, snapshot, document, null, options);
    }
    const record = await this.currentRevision(opened, documentId, revision);
    const material = await this.materialFor(opened, record);
    const snapshot = await this.snapshotForMaterial(opened, material);
    return this.measureSnapshot(
      material.config,
      snapshot,
      material.document,
      record.id,
      options,
    );
  }

  async validate(
    documentId: string,
    revision?: RevisionId | "HEAD",
  ): Promise<ValidationResult> {
    const snapshot = await this.getDocument(documentId, revision);
    const measurement = await this.measure(documentId, revision);
    return validateLegalDocument(snapshot.document, {
      revision: snapshot.revision,
      rulePack: snapshot.documentConfig.rulePack,
      filingKind: snapshot.documentConfig.filingKind,
      measurement: serializableMeasurement(measurement),
    });
  }

  async getDraftGuidance(
    documentId: string,
    revision?: RevisionId | "HEAD",
  ): Promise<DraftGuidance> {
    const selected = revision ?? "HEAD";
    const snapshot = await this.getDocument(documentId, selected);
    if (!snapshot.revision)
      throw new AgentDocxError("REVISION_NOT_FOUND", "Draft guidance requires a revision");
    const measurement = await this.measure(documentId, snapshot.revision, {
      paragraphDiagnostics: true,
      sectionDiagnostics: true,
    });
    return {
      schemaVersion: 1,
      documentId,
      revision: snapshot.revision,
      baseRevision: snapshot.revision,
      workingTreeHash: measurement.workingTreeHash,
      items: (measurement.deterministic.paragraphs ?? []).map((paragraph) => ({
        blockId: snapshot.document.blocks[paragraph.index]?.id ?? `b_${"0".repeat(36)}`,
        pages: Array.from(
          { length: paragraph.endPage - paragraph.startPage + 1 },
          (_, index) => paragraph.startPage + index,
        ),
        sectionIndex: 0,
        overflowingPage: measurement.budget?.withinLimit === false,
        overflowingSection: false,
        oneLineReduction: paragraph.oneLineReduction !== null,
        minimumReduction: {
          twips: paragraph.oneLineReduction?.estimatedRemovalTwips ?? 0,
          lines: paragraph.oneLineReduction ? 1 : 0,
        },
        editableSourceRanges: paragraph.lastLineSourceRanges.map((range) => ({
          start: range.position.start.offset,
          end: range.position.end.offset,
        })),
        lastLine: {
          usedTwips: paragraph.lastLineUsedTwips,
          availableTwips: paragraph.lastLineAvailableTwips,
          remainingTwips: paragraph.lastLineUnusedTwips,
        },
        budgets: {
          pageLimit: measurement.budget?.limitPages ?? null,
          pagesRemaining: measurement.budget?.pagesRemaining ?? null,
          countedLineLimit: null,
          countedLinesRemaining: null,
        },
      })),
    };
  }

  async evaluatePatch(
    patch: SourcePatch,
    options: { renderer?: "deterministic" | "word" | "libreoffice" | "compare" } = {},
  ): Promise<PatchEvaluation> {
    const opened = await openStore(this.manifestPath);
    const record = await this.currentRevision(opened, patch.documentId, patch.baseRevision);
    const material = await this.materialFor(opened, record);
    const baseSnapshot = await this.snapshotForMaterial(opened, material);
    const config = documentById(opened.manifest, patch.documentId);
    const current = await snapshotProjectDocument(opened, config);
    const beforeMeasurement = await this.measureSnapshot(
      material.config,
      baseSnapshot,
      material.document,
      record.id,
      {
        renderer: options.renderer,
        paragraphDiagnostics: true,
        sectionDiagnostics: true,
      },
    );
    const beforeValidation = validateLegalDocument(material.document, {
      revision: record.id,
      rulePack: material.config.rulePack,
      filingKind: material.config.filingKind,
      measurement: serializableMeasurement(beforeMeasurement),
    });
    const state = {
      headMatchesBase: (await readHead(opened.storePath, patch.documentId)) === record.id,
      sourceMatchesBase: current.sourceObject === record.sourceObject,
      documentConfigMatchesBase:
        current.documentConfigObject === record.documentConfigObject,
      dependenciesMatchBase:
        canonicalJson(current.dependencyObjects) ===
        canonicalJson(record.dependencyObjects),
      baseWorkingTreeHash: record.workingTreeHash,
      workingTreeHash: current.workingTreeHash,
    };
    const patchHash = objectId(canonicalJson(patch));
    try {
      if (
        patch.schemaVersion !== 1 ||
        patch.documentId !== record.documentId ||
        patch.baseRevision !== record.id
      )
        throw new AgentDocxError(
          "PATCH_INVALID",
          "Patch identity does not match its selected base revision",
        );
      let candidate = material.source;
      let previousEnd = -1;
      for (const edit of patch.edits) {
        if (
          !Number.isInteger(edit.start) ||
          !Number.isInteger(edit.end) ||
          edit.start < 0 ||
          edit.end < edit.start ||
          edit.start < previousEnd ||
          !isUtf16Boundary(material.source, edit.start) ||
          !isUtf16Boundary(material.source, edit.end) ||
          material.source.slice(edit.start, edit.end) !== edit.expectedText
        )
          throw new AgentDocxError(
            "PATCH_INVALID",
            "Patch edits must be sorted, non-overlapping, code-point-safe, and current",
          );
        previousEnd = edit.end;
      }
      for (const edit of [...patch.edits].sort((left, right) => right.start - left.start))
        candidate = `${candidate.slice(0, edit.start)}${edit.replacement}${candidate.slice(edit.end)}`;
      const sourceObject = objectId(candidate);
      const candidateSnapshot: ProjectSnapshot = {
        ...baseSnapshot,
        source: candidate,
        sourceObject,
        workingTreeHash: canonicalObjectId({
          sourceObject,
          documentConfigObject: baseSnapshot.documentConfigObject,
          dependencyObjects: baseSnapshot.dependencyObjects,
        }),
      };
      const parsedCandidate = documentFor(
        candidate,
        material.config,
        candidateSnapshot,
        opened.manifest.projectId,
        material.annotations,
        true,
      );
      const candidateDocument = {
        ...parsedCandidate,
        annotations: rebaseOpenAnnotations(
          material.document,
          parsedCandidate,
          material.annotations,
        ),
      };
      const candidateMeasurement = await measureNormalizedDocument(
        lowerLegalDocument(candidateDocument),
        {
          renderer: options.renderer,
          profile: material.config.profile,
          filingKind: material.config.filingKind,
          fontSet: sourceFontSet(material.config, candidateSnapshot),
          paragraphDiagnostics: true,
          sectionDiagnostics: true,
        },
      );
      const candidateValidation = validateLegalDocument(candidateDocument, {
        revision: record.id,
        rulePack: material.config.rulePack,
        filingKind: material.config.filingKind,
        measurement: serializableMeasurement(candidateMeasurement),
      });
      const pageLimitExcess = (measurement: MeasurementResult) =>
        measurement.budget && !measurement.budget.withinLimit
          ? Math.max(0, -measurement.budget.pagesRemaining)
          : 0;
      const countedLineExcess = (measurement: MeasurementResult) =>
        material.config.rulePack === "cand-civil@2026-05-01"
          ? measurement.deterministic.visualLinesByPage.reduce(
              (total, lines) => total + Math.max(0, lines - 28),
              0,
            )
          : 0;
      const affected = candidateDocument.blocks
        .filter((block) =>
          patch.edits.some(
            (edit) =>
              edit.start <= block.position.end.offset &&
              edit.end >= block.position.start.offset,
          ),
        )
        .map((block) => {
          const diagnostic = candidateMeasurement.deterministic.paragraphs?.find(
            (paragraph) =>
              paragraph.position.start.offset === block.position.start.offset,
          );
          return {
            blockId: block.id,
            sourceRanges: [
              {
                start: block.position.start.offset,
                end: block.position.end.offset,
              },
            ],
            pageSpan: diagnostic
              ? diagnostic.endPage - diagnostic.startPage + 1
              : 0,
            lineCount: diagnostic?.visualLines ?? 0,
            lastLineUsedTwips: diagnostic?.lastLineUsedTwips ?? 0,
          };
        });
      return {
        schemaVersion: 1,
        documentId: patch.documentId,
        patchHash,
        baseRevision: record.id,
        before: {
          measurement: serializableMeasurement(beforeMeasurement),
          validation: beforeValidation,
        },
        candidate: {
          status: "ok",
          measurement: serializableMeasurement(candidateMeasurement),
          validation: candidateValidation,
          deltas: {
            pageCount: candidateMeasurement.pageCount - beforeMeasurement.pageCount,
            countedLines:
              candidateMeasurement.deterministic.totalVisualLines -
              beforeMeasurement.deterministic.totalVisualLines,
            pageLimitExcess:
              pageLimitExcess(candidateMeasurement) -
              pageLimitExcess(beforeMeasurement),
            countedLineExcess:
              countedLineExcess(candidateMeasurement) -
              countedLineExcess(beforeMeasurement),
            lastPageUsedTwips:
              (candidateMeasurement.deterministic.lastPage?.usedTwips ?? 0) -
              (beforeMeasurement.deterministic.lastPage?.usedTwips ?? 0),
            validationSummary: {
              pass:
                candidateValidation.summary.pass - beforeValidation.summary.pass,
              fail:
                candidateValidation.summary.fail - beforeValidation.summary.fail,
              unknown:
                candidateValidation.summary.unknown -
                beforeValidation.summary.unknown,
            },
            affected,
            sections: [],
          },
        },
        passesConstraints:
          candidateValidation.status === "pass" &&
          pageLimitExcess(candidateMeasurement) === 0 &&
          countedLineExcess(candidateMeasurement) === 0,
        canApply:
          state.headMatchesBase &&
          state.sourceMatchesBase &&
          state.documentConfigMatchesBase &&
          state.dependenciesMatchBase,
        state,
      };
    } catch (error) {
      return {
        schemaVersion: 1,
        documentId: patch.documentId,
        patchHash,
        baseRevision: record.id,
        before: {
          measurement: serializableMeasurement(beforeMeasurement),
          validation: beforeValidation,
        },
        candidate: { status: "invalid", error: mutationError(error) },
        passesConstraints: false,
        canApply: false,
        state,
      };
    }
  }

  async applyPatch(
    patch: SourcePatch,
    input: {
      patchHash: string;
      gate?: "report" | "not-worse" | "pass";
      author: Actor;
      message: string;
    },
  ): Promise<RevisionMutationResult> {
    const evaluation = await this.evaluatePatch(patch);
    if (evaluation.patchHash !== input.patchHash)
      throw new AgentDocxError("PATCH_MISMATCH", "Patch hash does not match evaluation");
    if (evaluation.candidate.status !== "ok" || !evaluation.canApply)
      throw new AgentDocxError("PATCH_INVALID", "Patch cannot be applied to the working copy");
    return withLockedStore(this.manifestPath, async (opened) => {
      const config = documentById(opened.manifest, patch.documentId);
      const record = await this.currentRevision(opened, patch.documentId, patch.baseRevision);
      const material = await this.materialFor(opened, record);
      const baseSnapshot = await this.snapshotForMaterial(opened, material);
      const snapshot = await snapshotProjectDocument(opened, config);
      if (
        (await readHead(opened.storePath, patch.documentId)) !== record.id ||
        snapshot.sourceObject !== record.sourceObject ||
        snapshot.documentConfigObject !== record.documentConfigObject ||
        canonicalJson(snapshot.dependencyObjects) !==
          canonicalJson(record.dependencyObjects)
      )
        throw new AgentDocxError(
          "WORKING_COPY_CONFLICT",
          "Working copy changed before patch application",
        );
      if (
        patch.schemaVersion !== 1 ||
        patch.documentId !== record.documentId ||
        patch.baseRevision !== record.id
      )
        throw new AgentDocxError("PATCH_INVALID", "Patch identity is invalid");
      let previousEnd = -1;
      for (const edit of patch.edits) {
        if (
          typeof edit.expectedText !== "string" ||
          typeof edit.replacement !== "string" ||
          !Number.isInteger(edit.start) ||
          !Number.isInteger(edit.end) ||
          edit.start < 0 ||
          edit.end < edit.start ||
          edit.start < previousEnd ||
          !isUtf16Boundary(snapshot.source, edit.start) ||
          !isUtf16Boundary(snapshot.source, edit.end) ||
          snapshot.source.slice(edit.start, edit.end) !== edit.expectedText
        )
          throw new AgentDocxError(
            "PATCH_INVALID",
            "Patch edits are invalid, stale, or split a Unicode code point",
          );
        previousEnd = edit.end;
      }
      let source = snapshot.source;
      for (const edit of [...patch.edits].sort((left, right) => right.start - left.start))
        source = `${source.slice(0, edit.start)}${edit.replacement}${source.slice(edit.end)}`;
      const sourceObject = objectId(source);
      const candidateSnapshot: ProjectSnapshot = {
        ...snapshot,
        source,
        sourceObject,
        workingTreeHash: canonicalObjectId({
          sourceObject,
          documentConfigObject: snapshot.documentConfigObject,
          dependencyObjects: snapshot.dependencyObjects,
        }),
      };
      const parsedDocument = documentFor(
        source,
        config,
        candidateSnapshot,
        opened.manifest.projectId,
        material.annotations,
        true,
      );
      const annotations = rebaseOpenAnnotations(
        material.document,
        parsedDocument,
        material.annotations,
      );
      const document = { ...parsedDocument, annotations };
      const [beforeMeasurement, candidateMeasurement] = await Promise.all([
        this.measureSnapshot(
          material.config,
          baseSnapshot,
          material.document,
          record.id,
        ),
        this.measureSnapshot(config, candidateSnapshot, document, record.id),
      ]);
      const [beforeValidation, candidateValidation] = [
        validateLegalDocument(material.document, {
          revision: record.id,
          rulePack: material.config.rulePack,
          filingKind: material.config.filingKind,
          measurement: serializableMeasurement(beforeMeasurement),
        }),
        validateLegalDocument(document, {
          revision: record.id,
          rulePack: config.rulePack,
          filingKind: config.filingKind,
          measurement: serializableMeasurement(candidateMeasurement),
        }),
      ];
      const pageLimitExcess = (measurement: MeasurementResult) =>
        measurement.budget && !measurement.budget.withinLimit
          ? Math.max(0, -measurement.budget.pagesRemaining)
          : 0;
      const countedLineExcess = (measurement: MeasurementResult) =>
        config.rulePack === "cand-civil@2026-05-01"
          ? measurement.deterministic.visualLinesByPage.reduce(
              (total, lines) => total + Math.max(0, lines - 28),
              0,
            )
          : 0;
      const gate = input.gate ?? "not-worse";
      const worse =
        candidateValidation.summary.fail > beforeValidation.summary.fail ||
        candidateValidation.summary.unknown > beforeValidation.summary.unknown ||
        pageLimitExcess(candidateMeasurement) > pageLimitExcess(beforeMeasurement) ||
        countedLineExcess(candidateMeasurement) >
          countedLineExcess(beforeMeasurement);
      if (gate === "pass" && (
        candidateValidation.status !== "pass" ||
        pageLimitExcess(candidateMeasurement) !== 0 ||
        countedLineExcess(candidateMeasurement) !== 0
      ))
        throw new AgentDocxError("PATCH_FAILED_VALIDATION", "Patch does not pass constraints");
      if (gate === "not-worse" && worse)
        throw new AgentDocxError(
          "PATCH_FAILED_VALIDATION",
          "Patch worsens the selected revision's constraints",
        );
      await replaceOwnedFile(sourcePathFor(opened, config), snapshot.sourceObject, source);
      const refreshed = await snapshotProjectDocument(opened, config);
      return this.commitLocked(
        opened,
        config,
        refreshed,
        document,
        annotations,
        record.id,
        input.author,
        input.message,
        undefined,
        true,
      );
    });
  }

  async restore(
    documentId: string,
    input: {
      baseRevision: RevisionId | "HEAD";
      targetRevision: RevisionId | "HEAD";
      author: Actor;
      message: string;
    },
  ): Promise<RevisionMutationResult> {
    return withLockedStore(this.manifestPath, async (opened) => {
      const base = await this.currentRevision(opened, documentId, input.baseRevision);
      const target = await this.currentRevision(opened, documentId, input.targetRevision);
      if ((await readHead(opened.storePath, documentId)) !== base.id)
        throw new AgentDocxError("REVISION_CONFLICT", "Restore base is not the current head");
      const currentConfig = documentById(opened.manifest, documentId);
      const currentSnapshot = await snapshotProjectDocument(opened, currentConfig);
      if (currentSnapshot.workingTreeHash !== base.workingTreeHash)
        throw new AgentDocxError("WORKING_COPY_CONFLICT", "Working copy differs from restore base");
      const material = await this.materialFor(opened, target);
      await replaceOwnedFile(
        sourcePathFor(opened, currentConfig),
        currentSnapshot.sourceObject,
        material.source,
      );
      const manifest: AgentDocxManifest = {
        ...opened.manifest,
        documents: opened.manifest.documents.map((document) =>
          document.id === documentId ? material.config : document,
        ),
      };
      const updated = await updateManifest(opened, manifest);
      const snapshot = await snapshotProjectDocument(updated, material.config);
      return this.commitLocked(
        updated,
        material.config,
        snapshot,
        material.document,
        material.annotations,
        base.id,
        input.author,
        input.message,
        undefined,
        true,
      );
    });
  }

  async resolveChanges(
    documentId: string,
    input: ResolveChangesInput,
  ): Promise<RevisionMutationResult> {
    return withLockedStore(this.manifestPath, async (opened) => {
      const base = await this.currentRevision(
        opened,
        documentId,
        input.changeSet.baseRevision,
      );
      const head = await this.currentRevision(
        opened,
        documentId,
        input.changeSet.headRevision,
      );
      if (
        base.id === head.id ||
        (await readHead(opened.storePath, documentId)) !== head.id
      )
        throw new AgentDocxError(
          "REVISION_CONFLICT",
          "Change-set head must be the distinct current document head",
        );
      const baseMaterial = await this.materialFor(opened, base);
      const headMaterial = await this.materialFor(opened, head);
      const expected = createChangeSet(
        documentId,
        base.id,
        head.id,
        baseMaterial.document,
        headMaterial.document,
        baseMaterial.annotations,
        headMaterial.annotations,
        defaultAttribution(head.author, head.createdAt),
      );
      if (canonicalJson(expected) !== canonicalJson(input.changeSet))
        throw new AgentDocxError(
          "CHANGESET_INVALID",
          "Change set does not match the selected immutable revisions",
        );
      const changeIds = [
        ...expected.changes.map((change) => change.id),
        ...expected.annotations.map((change) => change.id),
      ].sort();
      const decisionIds = Object.keys(input.decisions).sort();
      if (
        changeIds.length !== decisionIds.length ||
        changeIds.some((id, index) => id !== decisionIds[index])
      )
        throw new AgentDocxError(
          "CHANGESET_INVALID",
          "Change-set decisions must select every change exactly once",
        );
      const config = documentById(opened.manifest, documentId);
      const snapshot = await snapshotProjectDocument(opened, config);
      if (
        snapshot.workingTreeHash !== head.workingTreeHash ||
        snapshot.sourceObject !== head.sourceObject ||
        snapshot.documentConfigObject !== head.documentConfigObject ||
        canonicalJson(snapshot.dependencyObjects) !== canonicalJson(head.dependencyObjects)
      )
        throw new AgentDocxError(
          "WORKING_COPY_CONFLICT",
          "Working copy differs from the change-set head",
        );
      const replacements = rejectedSourceReplacements(
        snapshot.source,
        expected.changes,
        input.decisions,
      );
      let source = snapshot.source;
      for (const replacement of replacements)
        source = `${source.slice(0, replacement.start)}${replacement.replacement}${source.slice(replacement.end)}`;
      let annotations = [...headMaterial.annotations];
      for (const change of expected.annotations) {
        if (input.decisions[change.id] !== "reject") continue;
        if (change.kind === "add")
          annotations = annotations.filter(
            (annotation) => annotation.id !== change.newValue.id,
          );
        else if (change.kind === "remove") annotations.push(change.oldValue);
        else
          annotations = annotations.map((annotation) =>
            annotation.id === change.newValue.id ? change.oldValue : annotation,
          );
      }
      if (source !== snapshot.source)
        await replaceOwnedFile(
          sourcePathFor(opened, config),
          snapshot.sourceObject,
          source,
        );
      const refreshed = source === snapshot.source
        ? snapshot
        : await snapshotProjectDocument(opened, config);
      const document = documentFor(
        refreshed.source,
        config,
        refreshed,
        opened.manifest.projectId,
        annotations,
        true,
      );
      return this.commitLocked(
        opened,
        config,
        refreshed,
        document,
        annotations,
        head.id,
        input.author,
        input.message,
        { schemaVersion: 1, changeSet: expected, decisions: input.decisions },
      );
    });
  }

  async exportDocx(
    documentId: string,
    input: ExportDocxInput,
  ): Promise<CompiledDocx> {
    return withLockedStore(this.manifestPath, async (opened) => {
      const record = await this.currentRevision(opened, documentId, input.revision);
      const material = await this.materialFor(opened, record);
      const snapshot = await this.snapshotForMaterial(opened, material);
      const fontSet = sourceFontSet(material.config, snapshot);
      const generationDependencies = new Map(
        [...snapshot.dependencyBytes].map(([key, dependency]) => {
          const sha256 = record.dependencyObjects[key];
          if (!sha256)
            throw new AgentDocxError(
              "PROJECT_INVALID",
              `Missing stored dependency hash: ${key}`,
            );
          return [key, { ...dependency, sha256 }] as const;
        }),
      );
      const compiled = await compileMarkdown(
        material.source,
        {
          projectId: opened.manifest.projectId,
          documentId,
          profile: material.config.profile,
          filingKind: material.config.filingKind,
          rulePack: material.config.rulePack,
          metadata: material.config.metadata,
          chrome: material.config.chrome,
          ...(snapshot.dependencyBytes.get("template")
            ? { template: snapshot.dependencyBytes.get("template")!.bytes }
            : {}),
          ...(fontSet ? { fontSet } : {}),
          assets: sourceAssets(snapshot),
        },
        {
          ...input.options,
          generation: {
            revision: record,
            annotations: input.mode === "redline" ? material.annotations : [],
            dependencies: generationDependencies,
          },
        },
      );
      let bytes = compiled.bytes;
      let mode: "clean" | "redline" = "clean";
      let baseRevision: RevisionId | null = null;
      let redlineVerification:
        | { revisionCount: number; commentCount: number; fieldCount: number }
        | undefined;
      if (input.mode === "redline") {
        const base = await this.currentRevision(
          opened,
          documentId,
          input.baseRevision,
        );
        if (
          base.id === record.id ||
          !(await this.isFirstParentAncestor(opened, base.id, record))
        )
          throw new AgentDocxError(
            "DOCX_REDLINE_UNSUPPORTED",
            "Redline base must be a distinct first-parent ancestor",
          );
        if (base.documentConfigObject !== record.documentConfigObject)
          throw new AgentDocxError(
            "DOCX_REDLINE_UNSUPPORTED",
            "Redline export does not support document configuration changes",
          );
        const baseMaterial = await this.materialFor(opened, base);
        const changeSet = createChangeSet(
          documentId,
          base.id,
          record.id,
          baseMaterial.document,
          material.document,
          baseMaterial.annotations,
          material.annotations,
          defaultAttribution(record.author, record.createdAt),
        );
        try {
          const generated = await generateRedlineDocx(
            baseMaterial.document,
            material.document,
            changeSet,
            compiled.measurement.deterministic.profile,
            {
              chrome: material.config.chrome,
              metadata: material.config.metadata,
              pageCount: Math.max(1, compiled.measurement.deterministic.pageCount),
              semanticManifest: createSemanticManifest({
                document: material.document,
                source: material.source,
                mode: "redline",
                attachments: compiled.attachments?.manifest ?? null,
                revision: record.id,
                baseRevision: base.id,
                validation: compiled.validation,
                dependencies: generationDependencies,
                changeSet,
                annotations: material.annotations,
              }),
              createdAt: record.createdAt,
            },
          );
          bytes = generated.bytes;
          redlineVerification = {
            revisionCount: generated.revisionCount,
            commentCount: generated.commentCount,
            fieldCount: 0,
          };
        } catch (error) {
          throw new AgentDocxError(
            "DOCX_REDLINE_UNSUPPORTED",
            error instanceof Error ? error.message : String(error),
          );
        }
        mode = "redline";
        baseRevision = base.id;
      }

      const importedAttachments = compiled.attachments
        ? {
            files: compiled.attachments.files,
            manifest: compiled.attachments.manifest,
          }
        : undefined;
      const inspected = await inspectDocxMaterial(bytes, {
        ...(importedAttachments ? { attachments: importedAttachments } : {}),
      });
      if (
        inspected.result.fidelity.overall === "unsupported" ||
        !inspected.semantic ||
        inspected.semantic.projectId !== opened.manifest.projectId ||
        inspected.semantic.documentId !== documentId ||
        inspected.semantic.mode !== mode ||
        inspected.semantic.revision !== record.id ||
        inspected.semantic.baseRevision !== baseRevision
      )
        throw new AgentDocxError(
          "DOCX_IMPORT_UNSUPPORTED",
          "Generated DOCX failed strict semantic re-import validation",
        );

      const destination = await assertExportDestination(opened, material.config, input.output);
      const owner = this.randomUuid();
      const stagePath = `${destination.output}.agent-docx-${owner}.stage`;
      if (await pathExists(stagePath))
        throw new AgentDocxError("OUTPUT_EXISTS", `DOCX export stage already exists: ${stagePath}`);
      const emptyObject = objectId(new Uint8Array());
      const initialIntent: ExportIntent = {
        schemaVersion: 1,
        state: "preparing",
        projectId: opened.manifest.projectId,
        manifestPath: opened.manifestPath,
        owner,
        outputPath: destination.output,
        attachmentPath: null,
        stagePath,
        docxStagePath: resolve(stagePath, "document.docx"),
        attachmentStagePath: null,
        docxSha256: emptyObject,
        attachmentManifestSha256: null,
        artifactStorePath: objectStorePath(opened.storePath, emptyObject),
        attachmentStorePath: null,
      };
      await updateExportIntent(opened.projectDirectory, initialIntent);
      let artifactObject!: RevisionId;
      let artifactStorePath!: string;
      let attachmentStorePath: string | null = null;
      let attachmentPath: string | null = null;
      let path = destination.output;
      let publicationStarted = false;
      try {
        await createExportStage(
          stagePath,
          owner,
          opened.manifest.projectId,
          opened.manifestPath,
        );
        await writeExclusiveFile(resolve(stagePath, "document.docx"), bytes);
        if (compiled.attachments) {
          attachmentPath = destination.attachment;
          await writeAttachmentStage(stagePath, compiled.attachments);
        }
        artifactObject = await writeObject(opened.storePath, bytes);
        artifactStorePath = objectStorePath(opened.storePath, artifactObject);
        if (compiled.attachments) {
          const manifestObject = await writeObject(
            opened.storePath,
            canonicalJson(compiled.attachments.manifest),
          );
          for (const entry of compiled.attachments.manifest.entries)
            await writeObject(
              opened.storePath,
              compiled.attachments.files[entry.name]!.bytes,
            );
          attachmentStorePath = objectStorePath(opened.storePath, manifestObject);
        }
        await updateExportIntent(opened.projectDirectory, {
          ...initialIntent,
          state: "prepared",
          attachmentPath,
          attachmentStagePath: compiled.attachments
            ? resolve(stagePath, "attachments")
            : null,
          docxSha256: artifactObject,
          attachmentManifestSha256: compiled.attachments?.manifestSha256 ?? null,
          artifactStorePath,
          attachmentStorePath,
        });
        publicationStarted = true;
        await completeExportIntent(opened.projectDirectory, opened.manifestPath);
      } catch (error) {
        if (!publicationStarted) {
          await rm(stagePath, { recursive: true, force: true }).catch(() => {});
          await clearExportIntent(opened.projectDirectory).catch(() => {});
        }
        throw error;
      }
      const artifact = {
        ...compiled.artifact,
        byteLength: bytes.byteLength,
        sha256: artifactObject,
        rendererProvenance: {
          ...compiled.artifact.rendererProvenance,
          ...(redlineVerification ? { verification: redlineVerification } : {}),
        },
        provenanceSha256: canonicalObjectId({
          generator: "agent-docx",
          documentId,
          revision: record.id,
          mode,
          baseRevision,
          docxSha256: artifactObject,
          dependencies: record.dependencyObjects,
          attachmentManifestSha256: compiled.attachments?.manifestSha256 ?? null,
        }),
        path,
        storePath: artifactStorePath,
        attachments:
          compiled.attachments && attachmentPath && attachmentStorePath
            ? {
                path: attachmentPath,
                storePath: attachmentStorePath,
                manifestSha256: compiled.attachments.manifestSha256,
                manifest: compiled.attachments.manifest,
              }
            : null,
        revision: record.id,
        mode,
        baseRevision,
      };
      return {
        ...compiled,
        bytes,
        measurement: {
          ...compiled.measurement,
          documentId,
          revision: record.id,
          workingTreeHash: record.workingTreeHash,
          dependencyObjects: record.dependencyObjects,
        },
        artifact,
      } as CompiledDocx;
    });
  }

  async importDocx(
    input: Extract<ImportDocxInput, { inspectOnly: false }>,
  ): Promise<DocxImportResult> {
    const inspected = await inspectDocxMaterial(input.input, {
      ...(input.attachments ? { attachments: input.attachments } : {}),
    });
    if (inspected.result.fidelity.overall === "unsupported")
      throw new AgentDocxError(
        "DOCX_IMPORT_UNSUPPORTED",
        "DOCX contains constructs that cannot be faithfully imported",
      );
    return withLockedStore(this.manifestPath, async (opened) => {
      const config = documentById(opened.manifest, input.documentId);
      if (!inspected.semantic)
        throw new AgentDocxError(
          "DOCX_IMPORT_UNSUPPORTED",
          "Strict DOCX import requires an agent-docx semantic manifest",
        );
      if (
        inspected.semantic.projectId !== opened.manifest.projectId ||
        inspected.semantic.documentId !== config.id
      )
        throw new AgentDocxError(
          "DOCX_IMPORT_UNSUPPORTED",
          "DOCX semantic manifest belongs to a different project document",
        );
      const sourcePath = sourcePathFor(opened, config);
      if (resolve(input.output) !== sourcePath)
        throw new AgentDocxError(
          "PROJECT_INVALID",
          "Import output must be the configured project document source",
        );
      if ((await readHead(opened.storePath, input.documentId)) !== null)
        throw new AgentDocxError(
          "REVISION_CONFLICT",
          "Strict DOCX import requires an empty document history",
        );
      const snapshot = await snapshotProjectDocument(opened, config);
      if (snapshot.source.length !== 0)
        throw new AgentDocxError(
          "WORKING_COPY_CONFLICT",
          "Strict DOCX import requires a zero-byte configured source",
        );
      const assetDestinations = importedAssetDestinations(
        opened,
        config,
        inspected.assets,
      );
      for (const asset of assetDestinations) {
        try {
          await lstat(asset.path);
        } catch (error) {
          if ((error as NodeJS.ErrnoException).code === "ENOENT") continue;
          throw error;
        }
        throw new AgentDocxError(
          "WORKING_COPY_CONFLICT",
          `Strict DOCX import requires an absent asset target: ${asset.source}`,
        );
      }
      if (inspected.tracked) {
        const baseDocument = parseLegalMarkdown(inspected.tracked.baseSource, {
          projectId: opened.manifest.projectId,
          documentId: config.id,
          metadata: config.metadata,
          chrome: config.chrome,
          assets: inspected.assets,
          requireMarkers: true,
        }).document;
        const headDocument = parseLegalMarkdown(inspected.tracked.headSource, {
          projectId: opened.manifest.projectId,
          documentId: config.id,
          metadata: config.metadata,
          chrome: config.chrome,
          assets: inspected.assets,
          requireMarkers: true,
          annotations: inspected.result.recognized.annotations,
        }).document;
        if (
          canonicalJson(semanticDocumentProjection(headDocument)) !==
          canonicalJson(inspected.semantic.document)
        )
          throw new AgentDocxError(
            "DOCX_IMPORT_UNSUPPORTED",
            "DOCX redline head document does not match the configured project document",
          );
        const writtenAssets: string[] = [];
        let currentSnapshot = snapshot;
        let baseCommitted = false;
        try {
          for (const asset of assetDestinations) {
            await mkdir(dirname(asset.path), { recursive: true, mode: 0o700 });
            await writeExclusiveFile(asset.path, asset.bytes);
            writtenAssets.push(asset.path);
          }
          await replaceOwnedFile(
            sourcePath,
            currentSnapshot.sourceObject,
            inspected.tracked.baseSource,
          );
          currentSnapshot = await snapshotProjectDocument(opened, config);
          const base = await this.commitLocked(
            opened,
            config,
            currentSnapshot,
            baseDocument,
            [],
            null,
            input.author,
            input.message,
          );
          baseCommitted = true;
          await replaceOwnedFile(
            sourcePath,
            currentSnapshot.sourceObject,
            inspected.tracked.headSource,
          );
          currentSnapshot = await snapshotProjectDocument(opened, config);
          const head = await this.commitLocked(
            opened,
            config,
            currentSnapshot,
            headDocument,
            inspected.result.recognized.annotations,
            base.revision.id,
            input.author,
            input.message,
          );
          return {
            ...inspected.result,
            inspectOnly: false,
            mode: "tracked",
            output: sourcePath,
            sourceSha256: currentSnapshot.sourceObject,
            baseRevision: base.revision.id,
            headRevision: head.revision.id,
            revisions: [base.revision.id, head.revision.id],
            recognized: {
              ...inspected.result.recognized,
              blocks: headDocument.blocks,
              footnotes: headDocument.footnotes,
            },
          } as DocxImportResult;
        } catch (error) {
          if (currentSnapshot.source !== snapshot.source)
            await replaceOwnedFile(
              sourcePath,
              currentSnapshot.sourceObject,
              snapshot.source,
            );
          if (baseCommitted)
            await rm(resolve(opened.storePath, "refs", `${config.id}.json`), {
              force: true,
            });
          for (const path of writtenAssets.reverse())
            await rm(path, { force: true });
          throw error;
        }
      }
      const document = parseLegalMarkdown(inspected.source, {
        projectId: opened.manifest.projectId,
        documentId: config.id,
        metadata: config.metadata,
        chrome: config.chrome,
        assets: inspected.assets,
        requireMarkers: true,
        annotations: inspected.result.recognized.annotations,
      }).document;
      if (
        canonicalJson(semanticDocumentProjection(document)) !==
        canonicalJson(inspected.semantic.document)
      )
        throw new AgentDocxError(
          "DOCX_IMPORT_UNSUPPORTED",
          "DOCX semantic document does not match the configured project document",
        );
      let published = false;
      let refreshed: ProjectSnapshot | null = null;
      const writtenAssets: string[] = [];
      try {
        for (const asset of assetDestinations) {
          await mkdir(dirname(asset.path), { recursive: true, mode: 0o700 });
          await writeExclusiveFile(asset.path, asset.bytes);
          writtenAssets.push(asset.path);
        }
        await replaceOwnedFile(sourcePath, snapshot.sourceObject, inspected.source);
        published = true;
        refreshed = await snapshotProjectDocument(opened, config);
        const mutation = await this.commitLocked(
          opened,
          config,
          refreshed,
          document,
          inspected.result.recognized.annotations,
          null,
          input.author,
          input.message,
        );
        return {
          ...inspected.result,
          inspectOnly: false,
          mode: "clean",
          output: sourcePath,
          sourceSha256: refreshed.sourceObject,
          baseRevision: mutation.revision.id,
          headRevision: mutation.revision.id,
          revisions: [mutation.revision.id],
          recognized: {
            ...inspected.result.recognized,
            blocks: document.blocks,
            footnotes: document.footnotes,
          },
        } as DocxImportResult;
      } catch (error) {
        if (published && refreshed)
          await replaceOwnedFile(sourcePath, refreshed.sourceObject, snapshot.source);
        for (const path of writtenAssets.reverse())
          await rm(path, { force: true });
        throw error;
      }
    });
  }
}

export const openProject = async (
  manifestPath: string = resolve(process.cwd(), "agent-docx.json"),
  options: ProjectRuntimeOptions = {},
): Promise<AgentDocxProject> => {
  await openStore(manifestPath);
  return new Project(
    resolve(manifestPath),
    options.clock ?? (() => new Date()),
    options.randomUUID ?? systemRandomUuid,
  );
};

export const createProject = async (
  manifestPath: string,
  input: ProjectDocumentInput,
  options: ProjectRuntimeOptions = {},
): Promise<AgentDocxProject> => {
  const absoluteManifestPath = resolve(manifestPath);
  const projectDirectory = dirname(absoluteManifestPath);
  const config = await documentConfigFromInput(projectDirectory, input);
  const projectId = options.randomUUID?.() ?? systemRandomUuid();
  const manifest: AgentDocxManifest = {
    schemaVersion: 1,
    projectId,
    defaultDocument: config.id,
    storeDir: ".agent-docx",
    documents: [config],
  };
  await initializeStore(absoluteManifestPath, manifest);
  if (input.createSource) await createEmptySource(projectDirectory, input.source);
  await withLockedStore(absoluteManifestPath, async (opened) =>
    materializeSourceMarkers(opened, config),
  );
  return new Project(
    absoluteManifestPath,
    options.clock ?? (() => new Date()),
    options.randomUUID ?? systemRandomUuid,
  );
};
