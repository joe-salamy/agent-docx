import { randomUUID as systemRandomUuid } from "node:crypto";
import { lstat, mkdir, open, readFile, rm } from "node:fs/promises";
import {
  dirname,
  extname,
  isAbsolute,
  relative,
  resolve,
  sep,
} from "node:path";
import { measureNormalizedDocument } from "../renderers/index.js";
import {
  validateLegalDocument,
  type ValidationResult,
} from "../legal/rules.js";
import {
  insertMissingBlockMarkers,
  parseLegalMarkdown,
  type LegalAssetInput,
} from "../legal/parse.js";
import { lowerLegalDocument } from "../legal/lower.js";
import type {
  Actor,
  AddressableBlock,
  LegalDocument,
  RevisionId,
  ReviewAnnotation,
} from "../legal/model.js";
import { AgentDocxError, type MeasurementResult } from "../types.js";
import {
  createChangeSet,
  createRevisionDelta,
  defaultAttribution,
  reattributeChangeSet,
  reattributeVisibleText,
  visibleTextForBlock,
  rebaseOpenAnnotations,
  type ChangeSetProvenance,
  type JsonObject,
} from "../revisions/diff.js";
import type {
  AttributionSpan,
  Change,
  ChangeAttribution,
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
  assertNoSymlinkComponents,
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
  removeInitializedProject,
  removeOwnedFile,
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
  ProjectCompiledDocx,
  ExportDocxInput,
  GeneratedAttachmentBundle,
  ImportDocxInput,
  DocxImportResult,
} from "../docx/contracts.js";
import type {
  DraftGuidance,
  PatchEvaluation,
  SourcePatch,
} from "../draft/types.js";
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
type CommitOptions = {
  expectedWorkingTreeHash?: RevisionId;
  parentIds?: readonly RevisionId[];
  firstParent?: RevisionRecord;
};

const version = "0.1.0";

const sourcePathFor = (
  opened: OpenedStore,
  config: AgentDocxDocumentConfig,
): string => resolve(opened.projectDirectory, config.source);
const snapshotWithSource = (
  snapshot: ProjectSnapshot,
  source: string,
): ProjectSnapshot => {
  const sourceObject = objectId(source);
  return {
    ...snapshot,
    source,
    sourceObject,
    workingTreeHash: canonicalObjectId({
      sourceObject,
      documentConfigObject: snapshot.documentConfigObject,
      dependencyObjects: snapshot.dependencyObjects,
    }),
  };
};

const snapshotWithDependencies = async (
  opened: OpenedStore,
  snapshot: ProjectSnapshot,
  config: AgentDocxDocumentConfig,
  dependencyObjects: Readonly<Record<string, RevisionId>>,
): Promise<ProjectSnapshot> => {
  const dependencyBytes = new Map<
    string,
    { bytes: Uint8Array; mediaType: string }
  >();
  for (const [key, dependencyObject] of Object.entries(dependencyObjects))
    dependencyBytes.set(key, {
      bytes: await readObject(opened.storePath, dependencyObject),
      mediaType:
        snapshot.dependencyBytes.get(key)?.mediaType ?? storedMediaType(key),
    });
  const documentConfigObject = canonicalObjectId(config);
  return {
    ...snapshot,
    documentConfigObject,
    dependencyObjects,
    dependencyBytes,
    workingTreeHash: canonicalObjectId({
      sourceObject: snapshot.sourceObject,
      documentConfigObject,
      dependencyObjects,
    }),
  };
};
const documentById = (
  manifest: AgentDocxManifest,
  documentId: string,
): AgentDocxDocumentConfig => {
  const document = manifest.documents.find((entry) => entry.id === documentId);
  if (!document)
    throw new AgentDocxError(
      "DOCUMENT_NOT_FOUND",
      `Document not found: ${documentId}`,
    );
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
      source
        .split("/")
        .some((part) => part === "" || part === "." || part === "..")
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
): Promise<string> => {
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
  return marked;
};

const sourceFontSet = (
  config: AgentDocxDocumentConfig,
  snapshot: ProjectSnapshot,
) => {
  if (!config.fontSet) return undefined;
  const regular = snapshot.dependencyBytes.get("font/regular");
  if (!regular)
    throw new AgentDocxError(
      "PROJECT_INVALID",
      "Configured regular font is missing",
    );
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
  return key.startsWith("rule-source/")
    ? "text/plain"
    : "application/octet-stream";
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
  left === right ||
  left.startsWith(`${right}${sep}`) ||
  right.startsWith(`${left}${sep}`);

const artifactDirectoryFor = (
  storePath: string,
  revision: RevisionId,
  provenance: RevisionId,
): string =>
  resolve(
    storePath,
    "artifacts",
    revision.slice("sha256:".length),
    provenance.slice("sha256:".length),
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
    config.assetsDir
      ? resolve(opened.projectDirectory, config.assetsDir)
      : null,
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
  await assertNoSymlinkComponents(absoluteOutput, "DOCX output");
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
    throw new AgentDocxError(
      "OUTPUT_EXISTS",
      `DOCX output already exists: ${absoluteOutput}`,
    );
  if (await pathExists(attachment))
    throw new AgentDocxError(
      "OUTPUT_EXISTS",
      `Attachment bundle already exists: ${attachment}`,
    );
  return { output: absoluteOutput, attachment };
};

const assertRegularDirectory = async (
  path: string,
  label: string,
): Promise<void> => {
  let entry;
  try {
    entry = await lstat(path);
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === "ENOENT")
      throw new AgentDocxError(
        "INPUT_NOT_FOUND",
        `${label} does not exist: ${path}`,
      );
    throw error;
  }
  if (!entry.isDirectory() || entry.isSymbolicLink())
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `${label} is not a regular directory: ${path}`,
    );
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
    if (
      !entry.payloadPath.startsWith("files/") ||
      entry.payloadPath.includes("\\")
    )
      throw new AgentDocxError(
        "PROJECT_INVALID",
        `Invalid attachment payload path: ${entry.payloadPath}`,
      );
    const payload = bundle.files[entry.name];
    if (!payload)
      throw new AgentDocxError(
        "PROJECT_INVALID",
        `Missing attachment bytes: ${entry.name}`,
      );
    const destination = resolve(attachmentRoot, entry.payloadPath);
    const relativePayload = relative(attachmentRoot, destination);
    if (
      relativePayload === ".." ||
      relativePayload.startsWith(`..${sep}`) ||
      isAbsolute(relativePayload)
    )
      throw new AgentDocxError(
        "PROJECT_INVALID",
        "Attachment payload escapes its bundle",
      );
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
    return {
      code: error.code,
      message: error.message,
      ...(error.details ? { details: error.details } : {}),
    };
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

const sourceRangeWithMarker = (
  source: string,
  start: number,
  end: number,
  blockId: string,
): { start: number; end: number } => {
  const marker = `<!-- agent-docx:block id="${blockId}" -->`;
  const markerStart = source.lastIndexOf(marker, start);
  if (
    markerStart < 0 ||
    (markerStart !== start &&
      !/^\r?\n$/.test(source.slice(markerStart + marker.length, start)))
  )
    return { start, end };
  const nextMarker = source.indexOf('<!-- agent-docx:block id="', end);
  if (nextMarker >= 0 && /^[\r\n]*$/.test(source.slice(end, nextMarker)))
    return { start: markerStart, end: nextMarker };
  const newlineLength = source.startsWith("\r\n", end)
    ? 2
    : source[end] === "\n"
      ? 1
      : 0;
  return {
    start: markerStart,
    end: end + newlineLength,
  };
};

type SourceMarkerLine = { id: string; start: number; end: number };
const sourceMarkerLines = (source: string): readonly SourceMarkerLine[] => {
  const markers: SourceMarkerLine[] = [];
  const pattern = /<!--[ \t]*agent-docx:block[ \t]+id="([^"]+)"[ \t]*-->/g;
  for (const match of source.matchAll(pattern)) {
    const markerStart = match.index ?? 0;
    const lineStart = source.lastIndexOf("\n", markerStart - 1) + 1;
    markers.push({
      id: match[1]!,
      start: lineStart,
      end: markerStart + match[0].length,
    });
  }
  return markers;
};
const sourceInsertionOffset = (
  baseSource: string,
  headSource: string,
  baseStart: number,
  baseEnd: number,
): number => {
  const baseMarkers = sourceMarkerLines(baseSource);
  const headMarkers = new Map(
    sourceMarkerLines(headSource).map((marker) => [marker.id, marker]),
  );
  const previous = [...baseMarkers]
    .reverse()
    .find((marker) => marker.end <= baseStart);
  const next = baseMarkers.find((marker) => marker.start >= baseEnd);
  const headNext = next ? headMarkers.get(next.id) : undefined;
  if (headNext) return headNext.start;
  const headPrevious = previous ? headMarkers.get(previous.id) : undefined;
  if (headPrevious) {
    const following = sourceMarkerLines(headSource).find(
      (marker) => marker.start > headPrevious.start,
    );
    return following?.start ?? headSource.length;
  }
  return 0;
};
const sourceInsertionText = (
  baseSource: string,
  baseStart: number,
  baseEnd: number,
): string => {
  const oldText = baseSource.slice(baseStart, baseEnd);
  const markers = sourceMarkerLines(baseSource);
  const next = markers.find((marker) => marker.start >= baseEnd);
  if (next) return oldText + baseSource.slice(baseEnd, next.start);
  return oldText + baseSource.slice(baseEnd);
};
const rejectedSourceReplacements = (
  source: string,
  baseSource: string,
  changes: readonly Change[],
  decisions: Readonly<Record<`c_${string}`, "accept" | "reject">>,
): RawReplacement[] => {
  const sourceChanges = changes.filter(
    (change) =>
      change.kind !== "add-config" &&
      change.kind !== "remove-config" &&
      change.kind !== "replace-config" &&
      change.kind !== "add-dependency" &&
      change.kind !== "remove-dependency" &&
      change.kind !== "replace-dependency",
  );
  if (
    sourceChanges.length > 0 &&
    sourceChanges.every((change) => decisions[change.id] === "reject")
  )
    return [
      {
        start: 0,
        end: source.length,
        expectedText: source,
        replacement: baseSource,
      },
    ];
  const replacements: RawReplacement[] = [];
  for (const change of changes) {
    if (decisions[change.id] !== "reject") continue;
    if (change.kind === "insert-block" || change.kind === "insert-text") {
      const range =
        change.kind === "insert-block"
          ? sourceRangeWithMarker(
              source,
              change.newSource.start,
              change.newSource.end,
              change.blockId,
            )
          : change.newSource;
      replacements.push({
        start: range.start,
        end: range.end,
        expectedText: source.slice(range.start, range.end),
        replacement: "",
      });
      continue;
    }
    if (change.kind === "delete-block") {
      const start = sourceInsertionOffset(
        baseSource,
        source,
        change.oldSource.start,
        change.oldSource.end,
      );
      replacements.push({
        start,
        end: start,
        expectedText: "",
        replacement: sourceInsertionText(
          baseSource,
          change.oldSource.start,
          change.oldSource.end,
        ),
      });
      continue;
    }
    if (change.kind === "move-block") {
      const range = change.newSource;
      const insertion = sourceInsertionOffset(
        baseSource,
        source,
        change.oldSource.start,
        change.oldSource.end,
      );
      replacements.push({
        start: range.start,
        end: range.end,
        expectedText: range.text,
        replacement: "",
      });
      replacements.push({
        start: insertion,
        end: insertion,
        expectedText: "",
        replacement: sourceInsertionText(
          baseSource,
          change.oldSource.start,
          change.oldSource.end,
        ),
      });
      continue;
    }
    if (change.kind === "replace-container-shell") {
      if (
        change.oldShell.sourceRanges.length !==
        change.newShell.sourceRanges.length
      )
        throw new AgentDocxError(
          "CHANGESET_INVALID",
          "Container shell source ranges do not align",
        );
      for (const [index, range] of change.newShell.sourceRanges.entries()) {
        const oldRange = change.oldShell.sourceRanges[index]!;
        replacements.push({
          start: range.start,
          end: range.end,
          expectedText: range.text,
          replacement: oldRange.text,
        });
      }
      continue;
    }
    if (change.kind === "delete-text") {
      replacements.push({
        start: change.newOffset,
        end: change.newOffset,
        expectedText: "",
        replacement: change.oldSource.text,
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
    if (
      change.kind === "add-config" ||
      change.kind === "remove-config" ||
      change.kind === "replace-config" ||
      change.kind === "add-dependency" ||
      change.kind === "remove-dependency" ||
      change.kind === "replace-dependency"
    )
      continue;
    throw new AgentDocxError(
      "CHANGESET_INVALID",
      "Cannot safely reject a change without an exact head source range",
    );
  }
  for (const replacement of replacements)
    if (
      source.slice(replacement.start, replacement.end) !==
      replacement.expectedText
    )
      throw new AgentDocxError(
        "REVISION_CONFLICT",
        "Change-set head source no longer matches its recorded range",
      );
  const ordered = [...replacements].sort(
    (left, right) => right.start - left.start,
  );
  for (const [index, replacement] of ordered.entries()) {
    const next = ordered[index + 1];
    if (next && replacement.start < next.end && next.start < replacement.end)
      throw new AgentDocxError(
        "CHANGESET_INVALID",
        "Rejected source changes overlap",
      );
  }
  return ordered;
};
type MutableJsonObject = Record<string, unknown>;

const jsonPointerParts = (path: string): readonly string[] => {
  if (path === "") return [];
  if (!path.startsWith("/"))
    throw new AgentDocxError(
      "CHANGESET_INVALID",
      `Invalid configuration path: ${path}`,
    );
  return path
    .slice(1)
    .split("/")
    .map((part) => part.replaceAll("~1", "/").replaceAll("~0", "~"));
};

const configParent = (
  root: MutableJsonObject,
  path: string,
): { parent: MutableJsonObject; key: string } => {
  const parts = jsonPointerParts(path);
  if (parts.length === 0)
    throw new AgentDocxError(
      "CHANGESET_INVALID",
      "Root configuration replacement is not supported",
    );
  let parent: MutableJsonObject = root;
  for (const part of parts.slice(0, -1)) {
    const child = parent[part];
    if (child === null || typeof child !== "object" || Array.isArray(child))
      throw new AgentDocxError(
        "CHANGESET_INVALID",
        `Configuration path is missing: ${path}`,
      );
    parent = child as MutableJsonObject;
  }
  return { parent, key: parts.at(-1)! };
};

const applyRejectedConfigChanges = (
  head: AgentDocxDocumentConfig,
  changes: readonly Change[],
  decisions: Readonly<Record<`c_${string}`, "accept" | "reject">>,
): AgentDocxDocumentConfig => {
  const result = JSON.parse(canonicalJson(head)) as MutableJsonObject;
  for (const change of changes) {
    if (decisions[change.id] !== "reject") continue;
    if (change.kind === "add-config") {
      const { parent, key } = configParent(result, change.path);
      const current = parent[key];
      if (
        !Object.hasOwn(parent, key) ||
        canonicalJson(current) !== canonicalJson(change.newValue)
      )
        throw new AgentDocxError(
          "CHANGESET_INVALID",
          `Configuration add does not match head: ${change.path}`,
        );
      delete parent[key];
    } else if (change.kind === "remove-config") {
      const { parent, key } = configParent(result, change.path);
      if (Object.hasOwn(parent, key))
        throw new AgentDocxError(
          "CHANGESET_INVALID",
          `Configuration remove does not match head: ${change.path}`,
        );
      parent[key] = change.oldValue;
    } else if (change.kind === "replace-config") {
      const { parent, key } = configParent(result, change.path);
      const current = parent[key];
      if (
        !Object.hasOwn(parent, key) ||
        canonicalJson(current) !== canonicalJson(change.newValue)
      )
        throw new AgentDocxError(
          "CHANGESET_INVALID",
          `Configuration replacement does not match head: ${change.path}`,
        );
      parent[key] = change.oldValue;
    }
  }
  return result as AgentDocxDocumentConfig;
};

const applyRejectedDependencyChanges = (
  head: Readonly<Record<string, RevisionId>>,
  changes: readonly Change[],
  decisions: Readonly<Record<`c_${string}`, "accept" | "reject">>,
): Record<string, RevisionId> => {
  const result = { ...head };
  for (const change of changes) {
    if (decisions[change.id] !== "reject") continue;
    if (change.kind === "add-dependency") {
      if (result[change.key] !== change.newObject)
        throw new AgentDocxError(
          "CHANGESET_INVALID",
          `Dependency add does not match head: ${change.key}`,
        );
      delete result[change.key];
    } else if (change.kind === "remove-dependency") {
      if (result[change.key] !== undefined)
        throw new AgentDocxError(
          "CHANGESET_INVALID",
          `Dependency remove does not match head: ${change.key}`,
        );
      result[change.key] = change.oldObject;
    } else if (change.kind === "replace-dependency") {
      if (result[change.key] !== change.newObject)
        throw new AgentDocxError(
          "CHANGESET_INVALID",
          `Dependency replacement does not match head: ${change.key}`,
        );
      result[change.key] = change.oldObject;
    }
  }
  return result;
};

const dependencyPath = (
  projectDirectory: string,
  config: AgentDocxDocumentConfig,
  key: string,
): string | null => {
  if (key === "template")
    return config.template ? resolve(projectDirectory, config.template) : null;
  if (key.startsWith("asset/")) {
    if (!config.assetsDir) return null;
    return resolve(
      projectDirectory,
      config.assetsDir,
      key.slice("asset/".length),
    );
  }
  if (key.startsWith("font/")) {
    const role = key.slice("font/".length) as
      | "regular"
      | "bold"
      | "italic"
      | "boldItalic";
    const configured = {
      regular: config.fontSet?.regularPath,
      bold: config.fontSet?.boldPath,
      italic: config.fontSet?.italicPath,
      boldItalic: config.fontSet?.boldItalicPath,
    }[role];
    return configured ? resolve(projectDirectory, configured) : null;
  }
  return null;
};
const dependencyPathsChanged = (
  projectDirectory: string,
  currentConfig: AgentDocxDocumentConfig,
  targetConfig: AgentDocxDocumentConfig,
  currentDependencies: Readonly<Record<string, RevisionId>>,
  targetDependencies: Readonly<Record<string, RevisionId>>,
): boolean =>
  [
    ...new Set([
      ...Object.keys(currentDependencies),
      ...Object.keys(targetDependencies),
    ]),
  ].some(
    (key) =>
      dependencyPath(projectDirectory, currentConfig, key) !==
      dependencyPath(projectDirectory, targetConfig, key),
  );

const materializeSelectedDependencies = async (
  opened: OpenedStore,
  currentConfig: AgentDocxDocumentConfig,
  targetConfig: AgentDocxDocumentConfig,
  currentDependencies: Readonly<Record<string, RevisionId>>,
  targetDependencies: Readonly<Record<string, RevisionId>>,
): Promise<void> => {
  const keys = [
    ...new Set([
      ...Object.keys(currentDependencies),
      ...Object.keys(targetDependencies),
    ]),
  ].sort((left, right) => left.localeCompare(right));
  const targetPaths = new Map<string, RevisionId>();
  const expectedIds = new Map<string, Set<RevisionId>>();
  const expectedId = (
    path: string | null,
    id: RevisionId | undefined,
  ): void => {
    if (path === null || id === undefined) return;
    const ids = expectedIds.get(path) ?? new Set<RevisionId>();
    ids.add(id);
    expectedIds.set(path, ids);
  };
  for (const key of keys) {
    const currentPath = dependencyPath(
      opened.projectDirectory,
      currentConfig,
      key,
    );
    const targetPath = dependencyPath(
      opened.projectDirectory,
      targetConfig,
      key,
    );
    expectedId(currentPath, currentDependencies[key]);
    expectedId(targetPath, targetDependencies[key]);
    const targetObject = targetDependencies[key];
    if (targetObject !== undefined && targetPath !== null) {
      const previous = targetPaths.get(targetPath);
      if (previous !== undefined && previous !== targetObject)
        throw new AgentDocxError(
          "PROJECT_INVALID",
          `Multiple dependencies target the same path: ${targetPath}`,
        );
      targetPaths.set(targetPath, targetObject);
    } else if (
      targetObject !== undefined &&
      key !== "profile" &&
      key !== "rule-pack" &&
      !key.startsWith("rule-source/")
    ) {
      throw new AgentDocxError(
        "PROJECT_INVALID",
        `Dependency has no configured path: ${key}`,
      );
    }
  }
  const paths = [
    ...new Set([...expectedIds.keys(), ...targetPaths.keys()]),
  ].sort((left, right) => left.localeCompare(right));
  const states = new Map<
    string,
    { bytes: Uint8Array | null; id: RevisionId | null }
  >();
  const knownIds = new Set<RevisionId>([
    ...Object.values(currentDependencies),
    ...Object.values(targetDependencies),
  ]);
  for (const path of paths) {
    try {
      const entry = await lstat(path);
      if (!entry.isFile() || entry.isSymbolicLink())
        throw new AgentDocxError(
          "WORKING_COPY_CONFLICT",
          `Owned dependency is not a regular file: ${path}`,
        );
      const bytes = await readFile(path);
      const id = objectId(bytes);
      if (!expectedIds.get(path)?.has(id))
        throw new AgentDocxError(
          "WORKING_COPY_CONFLICT",
          `Owned dependency changed: ${path}`,
        );
      states.set(path, { bytes, id });
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code !== "ENOENT") throw error;
      states.set(path, { bytes: null, id: null });
    }
  }
  const restore = async (): Promise<void> => {
    const restorePaths = [...paths].sort(
      (left, right) => right.length - left.length || left.localeCompare(right),
    );
    for (const path of restorePaths) {
      const original = states.get(path)!;
      let currentBytes: Uint8Array | null = null;
      try {
        const entry = await lstat(path);
        if (!entry.isFile() || entry.isSymbolicLink())
          throw new AgentDocxError(
            "PROJECT_INVALID",
            `Cannot roll back non-regular dependency: ${path}`,
          );
        currentBytes = await readFile(path);
      } catch (error) {
        if ((error as NodeJS.ErrnoException).code !== "ENOENT") throw error;
      }
      if (original.bytes === null) {
        if (currentBytes === null) continue;
        const currentId = objectId(currentBytes);
        if (!knownIds.has(currentId))
          throw new AgentDocxError(
            "WORKING_COPY_CONFLICT",
            `Dependency changed during rollback: ${path}`,
          );
        await removeOwnedFile(path, currentId);
        continue;
      }
      if (currentBytes === null) {
        await mkdir(dirname(path), { recursive: true, mode: 0o700 });
        await replaceOwnedFile(path, null, original.bytes);
        continue;
      }
      const currentId = objectId(currentBytes);
      if (currentId === original.id) continue;
      if (!knownIds.has(currentId))
        throw new AgentDocxError(
          "WORKING_COPY_CONFLICT",
          `Dependency changed during rollback: ${path}`,
        );
      await replaceOwnedFile(path, currentId, original.bytes);
    }
  };
  const removeCurrent = async (path: string): Promise<void> => {
    const state = states.get(path)!;
    if (state.id !== null) await removeOwnedFile(path, state.id);
  };
  const overlappingOldPaths = paths
    .filter(
      (path) =>
        !targetPaths.has(path) &&
        [...targetPaths.keys()].some((targetPath) =>
          pathsOverlap(path, targetPath),
        ),
    )
    .sort(
      (left, right) => right.length - left.length || left.localeCompare(right),
    );
  try {
    for (const path of overlappingOldPaths) await removeCurrent(path);
    for (const [targetPath, targetObject] of [...targetPaths.entries()].sort(
      ([left], [right]) => left.localeCompare(right),
    )) {
      const state = states.get(targetPath)!;
      if (state.id === targetObject) continue;
      const bytes = await readObject(opened.storePath, targetObject);
      await mkdir(dirname(targetPath), { recursive: true, mode: 0o700 });
      await replaceOwnedFile(targetPath, state.id, bytes);
    }
    for (const path of paths)
      if (!targetPaths.has(path) && !overlappingOldPaths.includes(path))
        await removeCurrent(path);
  } catch (error) {
    await restore();
    throw error;
  }
};
type AttributionState = {
  blocks: Map<string, readonly AttributionSpan[]>;
  operations: Map<string, ChangeAttribution>;
  config: Map<string, ChangeAttribution>;
  configOperations: Map<string, ChangeAttribution>;
  dependencies: Map<string, ChangeAttribution>;
  dependencyOperations: Map<string, ChangeAttribution>;
};

const provenanceBlocks = (
  document: LegalDocument,
): readonly AddressableBlock[] => {
  const result: AddressableBlock[] = [];
  const visit = (blocks: readonly AddressableBlock[]): void => {
    for (const block of blocks) {
      result.push(block);
      if (block.kind === "exhibit" || block.kind === "length-exclusion")
        visit(block.blocks);
      else if (block.kind === "list")
        for (const item of block.items) visit(item.children);
    }
  };
  visit(document.blocks);
  result.push(...document.footnotes);
  return result;
};
const textSpans = (
  text: string,
  attribution: ChangeAttribution,
): readonly AttributionSpan[] =>
  text.length === 0 ? [] : [{ start: 0, end: text.length, attribution }];

const seedAttributionState = (
  document: LegalDocument,
  config: AgentDocxDocumentConfig,
  attribution: ChangeAttribution,
): AttributionState => {
  const blocks = new Map<string, readonly AttributionSpan[]>();
  const operations = new Map<string, ChangeAttribution>();
  for (const block of provenanceBlocks(document)) {
    const spans = textSpans(visibleTextForBlock(block), attribution);
    blocks.set(block.id, spans);
    operations.set(block.id, attribution);
  }
  const configMap = new Map<string, ChangeAttribution>();
  const seedConfigAttribution = (
    value: unknown,
    path: string,
    valueAttribution: ChangeAttribution,
  ): void => {
    configMap.set(path, valueAttribution);
    if (value !== null && typeof value === "object" && !Array.isArray(value)) {
      for (const [key, child] of Object.entries(
        value as Record<string, unknown>,
      ))
        seedConfigAttribution(
          child,
          `${path}/${key.replaceAll("~", "~0").replaceAll("/", "~1")}`,
          valueAttribution,
        );
    }
  };
  seedConfigAttribution(config, "", attribution);
  const configOperations = new Map(
    [...configMap.keys()].map((path) => [path, attribution] as const),
  );
  return {
    blocks,
    operations,
    config: configMap,
    configOperations,
    dependencies: new Map(),
    dependencyOperations: new Map(),
  };
};
const setConfigAttribution = (
  config: Map<string, ChangeAttribution>,
  value: unknown,
  path: string,
  attribution: ChangeAttribution,
): void => {
  config.set(path, attribution);
  if (value !== null && typeof value === "object" && !Array.isArray(value)) {
    for (const [key, child] of Object.entries(value as Record<string, unknown>))
      setConfigAttribution(
        config,
        child,
        `${path}/${key.replaceAll("~", "~0").replaceAll("/", "~1")}`,
        attribution,
      );
  }
};
const removeConfigAttribution = (
  config: Map<string, ChangeAttribution>,
  path: string,
): void => {
  for (const key of config.keys())
    if (key === path || key.startsWith(`${path}/`)) config.delete(key);
};

const applyRevisionDelta = (
  state: AttributionState,
  changes: readonly Change[],
  previousDocument?: LegalDocument,
  currentDocument?: LegalDocument,
): void => {
  const textChanges = changes.filter(
    (
      change,
    ): change is Extract<
      Change,
      { kind: "insert-text" | "delete-text" | "replace-text" }
    > =>
      change.kind === "insert-text" ||
      change.kind === "delete-text" ||
      change.kind === "replace-text",
  );
  for (const change of changes) {
    if ("blockId" in change) {
      if (
        change.kind === "insert-text" ||
        change.kind === "delete-text" ||
        change.kind === "replace-text"
      )
        continue;
      if (change.kind === "delete-block") {
        state.blocks.delete(change.blockId);
        state.operations.set(change.blockId, change.attribution);
      } else if ("newAttributionSpans" in change) {
        state.blocks.set(change.blockId, change.newAttributionSpans);
        state.operations.set(change.blockId, change.attribution);
      }
      continue;
    }
    if (change.kind === "add-config" || change.kind === "replace-config") {
      removeConfigAttribution(state.config, change.path);
      setConfigAttribution(
        state.config,
        change.newValue,
        change.path,
        change.attribution,
      );
      state.configOperations.set(change.path, change.attribution);
    } else if (change.kind === "remove-config") {
      removeConfigAttribution(state.config, change.path);
      state.configOperations.set(change.path, change.attribution);
    } else if (
      change.kind === "add-dependency" ||
      change.kind === "replace-dependency"
    ) {
      state.dependencies.set(change.key, change.attribution);
      state.dependencyOperations.set(change.key, change.attribution);
    } else if (change.kind === "remove-dependency") {
      state.dependencies.delete(change.key);
      state.dependencyOperations.set(change.key, change.attribution);
    }
  }
  if (!previousDocument || !currentDocument) return;
  const previousBlocks = new Map(
    provenanceBlocks(previousDocument).map((block) => [block.id, block]),
  );
  const currentBlocks = new Map(
    provenanceBlocks(currentDocument).map((block) => [block.id, block]),
  );
  for (const blockId of new Set(textChanges.map((change) => change.blockId))) {
    const previousBlock = previousBlocks.get(blockId);
    const currentBlock = currentBlocks.get(blockId);
    if (!previousBlock || !currentBlock)
      throw new AgentDocxError(
        "PROJECT_INVALID",
        `Text delta references a missing block: ${blockId}`,
      );
    const operation = textChanges.at(-1)!;
    state.blocks.set(
      blockId,
      reattributeVisibleText(
        visibleTextForBlock(previousBlock),
        visibleTextForBlock(currentBlock),
        state.blocks.get(blockId),
        operation.attribution,
      ),
    );
    state.operations.set(blockId, operation.attribution);
  }
  const sourceChanges = changes.filter(
    (change): change is Extract<Change, { blockId: string }> =>
      "blockId" in change,
  );
  const sourceOperation = sourceChanges.at(-1);
  if (!sourceOperation) return;
  const parentMap = (
    document: LegalDocument,
  ): {
    blocks: Map<string, AddressableBlock>;
    parents: Map<string, string | null>;
  } => {
    const blocks = new Map<string, AddressableBlock>();
    const parents = new Map<string, string | null>();
    const visit = (
      entries: readonly AddressableBlock[],
      parent: string | null,
    ): void => {
      for (const block of entries) {
        blocks.set(block.id, block);
        parents.set(block.id, parent);
        if (block.kind === "exhibit" || block.kind === "length-exclusion")
          visit(block.blocks, block.id);
        else if (block.kind === "list")
          for (const item of block.items) visit(item.children, block.id);
      }
    };
    visit(document.blocks, null);
    for (const footnote of document.footnotes) {
      blocks.set(footnote.id, footnote);
      parents.set(footnote.id, null);
    }
    return { blocks, parents };
  };
  const previousTree = parentMap(previousDocument);
  const currentTree = parentMap(currentDocument);
  const affectedContainers = new Set<string>();
  const addAncestors = (
    blockId: string,
    tree: ReturnType<typeof parentMap>,
  ): void => {
    let parentId = tree.parents.get(blockId) ?? null;
    while (parentId !== null) {
      const parent = tree.blocks.get(parentId);
      if (
        parent &&
        (parent.kind === "list" ||
          parent.kind === "exhibit" ||
          parent.kind === "length-exclusion")
      )
        affectedContainers.add(parentId);
      parentId = tree.parents.get(parentId) ?? null;
    }
  };
  for (const change of sourceChanges) {
    addAncestors(change.blockId, previousTree);
    addAncestors(change.blockId, currentTree);
  }
  for (const blockId of affectedContainers) {
    const previousBlock = previousTree.blocks.get(blockId);
    const currentBlock = currentTree.blocks.get(blockId);
    if (!previousBlock || !currentBlock) continue;
    state.blocks.set(
      blockId,
      reattributeVisibleText(
        visibleTextForBlock(previousBlock),
        visibleTextForBlock(currentBlock),
        state.blocks.get(blockId),
        sourceOperation.attribution,
      ),
    );
    state.operations.set(blockId, sourceOperation.attribution);
  }
};

const provenanceForRevision = async (
  opened: OpenedStore,
  target: RevisionRecord,
): Promise<AttributionState> => {
  const chain: RevisionRecord[] = [];
  const visited = new Set<RevisionId>();
  let current: RevisionRecord | null = target;
  while (current) {
    if (current.documentId !== target.documentId)
      throw new AgentDocxError(
        "PROJECT_INVALID",
        `Revision belongs to another document: ${current.id}`,
      );
    if (visited.has(current.id))
      throw new AgentDocxError(
        "PROJECT_INVALID",
        "Revision graph contains a cycle",
      );
    visited.add(current.id);
    chain.push(current);
    const parent: RevisionId | undefined = current.parents[0];
    current = parent
      ? await readRevisionJson<RevisionRecord>(opened.storePath, parent)
      : null;
  }
  chain.reverse();
  const root = chain[0]!;
  const rootMaterial = await (async () => {
    const config = JSON.parse(
      new TextDecoder("utf-8", { fatal: true }).decode(
        await readObject(opened.storePath, root.documentConfigObject),
      ),
    ) as AgentDocxDocumentConfig;
    const document = JSON.parse(
      new TextDecoder("utf-8", { fatal: true }).decode(
        await readObject(opened.storePath, root.legalDocumentObject),
      ),
    ) as LegalDocument;
    return { config, document };
  })();
  const state = seedAttributionState(
    rootMaterial.document,
    rootMaterial.config,
    defaultAttribution(root.author, root.createdAt),
  );
  for (const key of Object.keys(root.dependencyObjects)) {
    const attribution = defaultAttribution(root.author, root.createdAt);
    state.dependencies.set(key, attribution);
    state.dependencyOperations.set(key, attribution);
  }
  let previousDocument = rootMaterial.document;
  for (const [index, record] of chain.slice(1).entries()) {
    const parent = chain[index]!;
    if (
      record.documentId !== target.documentId ||
      record.parents[0] !== parent.id
    )
      throw new AgentDocxError(
        "PROJECT_INVALID",
        `Revision delta parent is invalid: ${record.id}`,
      );
    if (!record.deltaObject)
      throw new AgentDocxError(
        "PROJECT_INVALID",
        `Revision delta is missing: ${record.id}`,
      );
    const delta = JSON.parse(
      new TextDecoder("utf-8", { fatal: true }).decode(
        await readObject(opened.storePath, record.deltaObject),
      ),
    ) as {
      schemaVersion?: unknown;
      parentSourceObject?: unknown;
      parentDocumentConfigObject?: unknown;
      changes?: readonly Change[];
      annotations?: readonly unknown[];
    };
    if (
      delta.schemaVersion !== 1 ||
      delta.parentSourceObject !== parent.sourceObject ||
      delta.parentDocumentConfigObject !== parent.documentConfigObject ||
      !Array.isArray(delta.changes) ||
      !Array.isArray(delta.annotations)
    )
      throw new AgentDocxError(
        "PROJECT_INVALID",
        `Revision delta is malformed: ${record.id}`,
      );
    const currentDocument = JSON.parse(
      new TextDecoder("utf-8", { fatal: true }).decode(
        await readObject(opened.storePath, record.legalDocumentObject),
      ),
    ) as LegalDocument;
    applyRevisionDelta(state, delta.changes, previousDocument, currentDocument);
    previousDocument = currentDocument;
  }
  return state;
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
      throw new AgentDocxError(
        "REVISION_NOT_FOUND",
        `Document has no revision: ${documentId}`,
      );
    const visited = new Set<RevisionId>();
    const pending: {
      id: RevisionId;
      ancestry: ReadonlySet<RevisionId>;
    }[] = head ? [{ id: head, ancestry: new Set() }] : [];
    while (pending.length > 0) {
      const entry = pending.pop()!;
      if (entry.ancestry.has(entry.id))
        throw new AgentDocxError(
          "PROJECT_INVALID",
          "Revision graph contains a cycle",
        );
      if (visited.has(entry.id)) continue;
      visited.add(entry.id);
      const record = await readRevisionJson<RevisionRecord>(
        opened.storePath,
        entry.id,
      );
      if (record.documentId !== documentId)
        throw new AgentDocxError(
          "PROJECT_INVALID",
          `Revision belongs to another document: ${entry.id}`,
        );
      if (record.id === requested) return record;
      const ancestry = new Set(entry.ancestry);
      ancestry.add(entry.id);
      pending.push(...record.parents.map((id) => ({ id, ancestry })));
    }
    throw new AgentDocxError(
      "REVISION_NOT_FOUND",
      `Revision not found: ${requested}`,
    );
  }

  private async isFirstParentAncestor(
    opened: OpenedStore,
    ancestor: RevisionId,
    descendant: RevisionRecord,
  ): Promise<boolean> {
    const visited = new Set<RevisionId>();
    let current: RevisionRecord | null = descendant;
    while (current) {
      if (visited.has(current.id))
        throw new AgentDocxError(
          "PROJECT_INVALID",
          "Revision graph contains a cycle",
        );
      visited.add(current.id);
      if (current.documentId !== descendant.documentId)
        throw new AgentDocxError(
          "PROJECT_INVALID",
          "Revision graph crosses documents",
        );
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
    const record = await readRevisionJson<RevisionRecord>(
      opened.storePath,
      head,
    );
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
    const { includeGeneratedDocx, ...safeOptions } =
      options as ProjectMeasureOptions & {
        includeGeneratedDocx?: boolean;
      };
    void includeGeneratedDocx;
    const measurement = await measureNormalizedDocument(
      lowerLegalDocument(document),
      {
        ...safeOptions,
        profile: config.profile,
        filingKind: config.filingKind,
        chrome: config.chrome,
        fontSet: sourceFontSet(config, snapshot),
      },
    );
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
  ): Promise<{
    source: string;
    snapshot: ProjectSnapshot;
    document: LegalDocument;
  }> {
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
        snapshot = snapshotWithSource(snapshot, marked);
        source = marked;
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
    requireWorkingTreeMatch = false,
    options: CommitOptions = {},
  ): Promise<RevisionMutationResult> {
    const head = await readHead(opened.storePath, config.id);
    if (base === null ? head !== null : head !== base)
      throw new AgentDocxError(
        "REVISION_CONFLICT",
        "Revision base does not match current head",
      );
    const headParent = head
      ? await readRevisionJson<RevisionRecord>(opened.storePath, head)
      : null;
    const parent = options.firstParent ?? headParent;
    const parentMaterial = parent
      ? await this.materialFor(opened, parent)
      : null;
    const headParentMaterial =
      headParent && headParent.id === parent?.id
        ? parentMaterial
        : headParent
          ? await this.materialFor(opened, headParent)
          : null;
    const parentProvenance = parent
      ? await provenanceForRevision(opened, parent)
      : null;
    const committedAnnotations =
      headParentMaterial && !preserveAnnotations
        ? rebaseOpenAnnotations(
            headParentMaterial.document,
            document,
            annotations,
          )
        : annotations;
    const committedDocument = {
      ...document,
      annotations: committedAnnotations,
    };
    await storeSnapshot(opened, snapshot);
    const sourceObject = await writeObject(opened.storePath, snapshot.source);
    const documentConfigObject = await writeObject(
      opened.storePath,
      canonicalJson(config),
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
              parentMaterial!.config as unknown as JsonObject,
              config as unknown as JsonObject,
              snapshot.dependencyObjects,
              parentMaterial!.source,
              snapshot.source,
              parentProvenance!.blocks,
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
      parents: options.parentIds ?? (headParent ? [headParent.id] : []),
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
    const currentConfig = documentById(opened.manifest, config.id);
    const currentSnapshot = await snapshotProjectDocument(
      opened,
      currentConfig,
    );
    if (
      requireWorkingTreeMatch &&
      options.expectedWorkingTreeHash !== undefined &&
      currentSnapshot.workingTreeHash !== options.expectedWorkingTreeHash
    )
      throw new AgentDocxError(
        "WORKING_COPY_CONFLICT",
        "Working copy changed during the mutation",
      );
    let sourceChanged = false;
    let manifestChanged = false;
    let headChanged = false;
    let dependenciesChanged = false;
    try {
      dependenciesChanged =
        dependencyPathsChanged(
          opened.projectDirectory,
          currentConfig,
          config,
          currentSnapshot.dependencyObjects,
          snapshot.dependencyObjects,
        ) ||
        canonicalJson(currentSnapshot.dependencyObjects) !==
          canonicalJson(snapshot.dependencyObjects);
      if (dependenciesChanged)
        await materializeSelectedDependencies(
          opened,
          currentConfig,
          config,
          currentSnapshot.dependencyObjects,
          snapshot.dependencyObjects,
        );
      if (currentSnapshot.sourceObject !== snapshot.sourceObject) {
        await replaceOwnedFile(
          sourcePathFor(opened, currentConfig),
          currentSnapshot.sourceObject,
          snapshot.source,
        );
        sourceChanged = true;
      }
      const targetManifest: AgentDocxManifest = {
        ...opened.manifest,
        documents: opened.manifest.documents.map((document) =>
          document.id === config.id ? config : document,
        ),
      };
      if (canonicalJson(targetManifest) !== canonicalJson(opened.manifest)) {
        await updateManifest(opened, targetManifest);
        manifestChanged = true;
      }
      await writeHead(opened.storePath, config.id, revisionId);
      headChanged = true;
    } catch (error) {
      if (headChanged) {
        if (head === null)
          await rm(resolve(opened.storePath, "refs", `${config.id}.json`), {
            force: true,
          });
        else await writeHead(opened.storePath, config.id, head);
      }
      if (dependenciesChanged)
        await materializeSelectedDependencies(
          opened,
          config,
          currentConfig,
          snapshot.dependencyObjects,
          currentSnapshot.dependencyObjects,
        );
      if (manifestChanged) await updateManifest(opened, opened.manifest);
      if (sourceChanged)
        await replaceOwnedFile(
          sourcePathFor(opened, currentConfig),
          snapshot.sourceObject,
          currentSnapshot.source,
        );
      throw error;
    }
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
            documentConfig:
              record?.documentConfigObject === snapshot.documentConfigObject,
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
      if (
        opened.manifest.documents.some(
          (document) => document.id === input.documentId,
        )
      )
        throw new AgentDocxError(
          "DOCUMENT_EXISTS",
          `Document already exists: ${input.documentId}`,
        );
      const config = await documentConfigFromInput(
        opened.projectDirectory,
        input,
      );
      const sourcePath = sourcePathFor(opened, config);
      const originalSource = input.createSource
        ? null
        : await readFile(sourcePath);
      if (input.createSource)
        await createEmptySource(opened.projectDirectory, input.source);
      const manifest: AgentDocxManifest = {
        ...opened.manifest,
        defaultDocument: input.makeDefault
          ? config.id
          : opened.manifest.defaultDocument,
        documents: [...opened.manifest.documents, config],
      };
      try {
        await materializeSourceMarkers(opened, config);
        await updateManifest(opened, manifest);
      } catch (error) {
        const snapshot = await snapshotProjectDocument(opened, config);
        if (originalSource === null)
          await removeOwnedFile(sourcePath, snapshot.sourceObject);
        else
          await replaceOwnedFile(
            sourcePath,
            snapshot.sourceObject,
            originalSource,
          );
        throw error;
      }
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
        throw new AgentDocxError(
          "REVISION_CONFLICT",
          "Document configuration base does not match head",
        );
      const currentSnapshot = await snapshotProjectDocument(opened, current);
      if (head !== null) {
        const headRecord = await readRevisionJson<RevisionRecord>(
          opened.storePath,
          head,
        );
        if (currentSnapshot.workingTreeHash !== headRecord.workingTreeHash)
          throw new AgentDocxError(
            "WORKING_COPY_CONFLICT",
            "Working copy differs from the configuration head",
          );
      }
      if (Object.keys(input.changes).length === 0)
        throw new AgentDocxError(
          "INVALID_ARGUMENT",
          "Document configuration changes are required",
        );
      const next: AgentDocxDocumentConfig = { ...current };
      if (input.changes.profile !== undefined)
        next.profile = input.changes.profile;
      if (input.changes.metadata !== undefined)
        next.metadata = input.changes.metadata;
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
      const snapshot = await snapshotProjectDocument(opened, next);
      const annotations = await this.annotationsForHead(opened, documentId);
      const prepared = await this.prepareWorkingDocument(
        opened,
        next,
        snapshot,
        annotations,
        true,
      );
      return this.commitLocked(
        opened,
        next,
        prepared.snapshot,
        prepared.document,
        annotations,
        base,
        input.author,
        input.message,
        undefined,
        false,
        true,
        { expectedWorkingTreeHash: currentSnapshot.workingTreeHash },
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
    const headAnnotations = await this.annotationsForHead(opened, documentId);
    let document = documentFor(
      snapshot.source,
      config,
      snapshot,
      opened.manifest.projectId,
      [],
      false,
    );
    let annotations = headAnnotations;
    if (head) {
      const headRecord = await this.currentRevision(opened, documentId, head);
      const headMaterial = await this.materialFor(opened, headRecord);
      annotations = rebaseOpenAnnotations(
        headMaterial.document,
        document,
        headAnnotations,
      );
      document = { ...document, annotations };
    }
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
    input: {
      baseRevision: RevisionId | "HEAD" | null;
      author: Actor;
      message: string;
    },
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
        undefined,
        false,
        true,
        { expectedWorkingTreeHash: snapshot.workingTreeHash },
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
    const pending: {
      id: RevisionId;
      ancestry: ReadonlySet<RevisionId>;
    }[] = head ? [{ id: head, ancestry: new Set() }] : [];
    while (pending.length > 0) {
      const entry = pending.pop()!;
      if (entry.ancestry.has(entry.id))
        throw new AgentDocxError(
          "PROJECT_INVALID",
          "Revision graph contains a cycle",
        );
      if (reachable.has(entry.id)) continue;
      const record = await readRevisionJson<RevisionRecord>(
        opened.storePath,
        entry.id,
      );
      if (record.documentId !== documentId)
        throw new AgentDocxError(
          "PROJECT_INVALID",
          "Revision graph crosses documents",
        );
      reachable.set(entry.id, record);
      const ancestry = new Set(entry.ancestry);
      ancestry.add(entry.id);
      pending.push(...record.parents.map((id) => ({ id, ancestry })));
    }
    const ordered = [...reachable.values()].sort(
      (left, right) =>
        right.createdAt.localeCompare(left.createdAt) ||
        left.id.localeCompare(right.id),
    );
    const limit = input.limit ?? 100;
    if (!Number.isInteger(limit) || limit < 1 || limit > 1000)
      throw new AgentDocxError(
        "INVALID_ARGUMENT",
        "Revision limit must be 1 through 1000",
      );
    const start = input.cursor
      ? Math.max(
          0,
          ordered.findIndex((record) => record.id === input.cursor) + 1,
        )
      : 0;
    if (input.cursor && start === 0)
      throw new AgentDocxError(
        "REVISION_NOT_FOUND",
        `Revision cursor not found: ${input.cursor}`,
      );
    const items = ordered.slice(start, start + limit);
    return {
      schemaVersion: 1,
      items,
      nextCursor: ordered[start + limit]?.id ?? null,
    };
  }

  async getRevision(
    documentId: string,
    revision: RevisionId | "HEAD",
  ): Promise<RevisionRecord> {
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
    if (
      baseRecord.id !== headRecord.id &&
      !(await this.isFirstParentAncestor(opened, baseRecord.id, headRecord))
    )
      throw new AgentDocxError(
        "REVISION_CONFLICT",
        "Diff base must be a first-parent ancestor of head",
      );
    const baseMaterial = await this.materialFor(opened, baseRecord);
    const headMaterial = await this.materialFor(opened, headRecord);
    const changeSet = createChangeSet(
      documentId,
      baseRecord.id,
      headRecord.id,
      baseMaterial.document,
      headMaterial.document,
      baseMaterial.annotations,
      headMaterial.annotations,
      defaultAttribution(headRecord.author, headRecord.createdAt),
      {
        baseConfig: baseMaterial.config as unknown as JsonObject,
        headConfig: headMaterial.config as unknown as JsonObject,
        baseDependencies: baseRecord.dependencyObjects,
        headDependencies: headRecord.dependencyObjects,
        baseSource: baseMaterial.source,
        headSource: headMaterial.source,
      },
    );
    const baseProvenance = await provenanceForRevision(opened, baseRecord);
    const headProvenance = await provenanceForRevision(opened, headRecord);
    const provenance: ChangeSetProvenance = {
      baseBlocks: baseProvenance.blocks,
      headBlocks: headProvenance.blocks,
      baseOperations: baseProvenance.operations,
      headOperations: headProvenance.operations,
      baseConfig: baseProvenance.config,
      headConfig: headProvenance.config,
      baseDependencies: baseProvenance.dependencies,
      headDependencies: headProvenance.dependencies,
      baseConfigOperations: baseProvenance.configOperations,
      headConfigOperations: headProvenance.configOperations,
      baseDependencyOperations: baseProvenance.dependencyOperations,
      headDependencyOperations: headProvenance.dependencyOperations,
      baseDocument: baseMaterial.document,
      headDocument: headMaterial.document,
    };
    return reattributeChangeSet(changeSet, provenance);
  }

  async addReview(
    documentId: string,
    input: AddReviewInput,
  ): Promise<RevisionMutationResult> {
    return withLockedStore(this.manifestPath, async (opened) => {
      const record = await this.currentRevision(
        opened,
        documentId,
        input.revision,
      );
      const head = await readHead(opened.storePath, documentId);
      if (head !== record.id)
        throw new AgentDocxError(
          "REVISION_CONFLICT",
          "Review must target the current head",
        );
      const material = await this.materialFor(opened, record);
      const block = [
        ...material.document.blocks,
        ...material.document.footnotes,
      ].find((entry) => entry.id === input.blockId);
      if (!block)
        throw new AgentDocxError(
          "REFERENCE_INVALID",
          `Block not found: ${input.blockId}`,
        );
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
        !["paragraph", "blockquote", "heading", "numbered-paragraph"].includes(
          block.kind,
        ) &&
        !(block.kind === "footnote" && block.paragraphs.length === 1)
      )
        throw new AgentDocxError(
          "ANNOTATION_CONFLICT",
          "Review ranges must stay within one source-mapped paragraph",
        );
      const annotationId = `a_${this.randomUuid()}` as `a_${string}`;
      if (
        !/^a_[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/.test(
          annotationId,
        )
      )
        throw new AgentDocxError(
          "PROJECT_INVALID",
          "Runtime randomUUID did not return a UUIDv4",
        );
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
        undefined,
        false,
        true,
        { expectedWorkingTreeHash: snapshot.workingTreeHash },
      );
    });
  }

  async resolveReview(
    documentId: string,
    input: ResolveReviewInput,
  ): Promise<RevisionMutationResult> {
    return withLockedStore(this.manifestPath, async (opened) => {
      const record = await this.currentRevision(
        opened,
        documentId,
        input.revision,
      );
      const head = await readHead(opened.storePath, documentId);
      if (head !== record.id)
        throw new AgentDocxError(
          "REVISION_CONFLICT",
          "Review must target the current head",
        );
      const material = await this.materialFor(opened, record);
      const annotations = material.annotations.map((annotation) =>
        annotation.id === input.annotationId
          ? { ...annotation, status: "resolved" as const }
          : annotation,
      );
      if (
        annotations.every((annotation) => annotation.id !== input.annotationId)
      )
        throw new AgentDocxError(
          "REFERENCE_INVALID",
          `Annotation not found: ${input.annotationId}`,
        );
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
        undefined,
        false,
        true,
        { expectedWorkingTreeHash: snapshot.workingTreeHash },
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
      const head = await readHead(opened.storePath, documentId);
      const headAnnotations = await this.annotationsForHead(opened, documentId);
      let document = documentFor(
        snapshot.source,
        config,
        snapshot,
        opened.manifest.projectId,
        [],
        false,
      );
      let annotations = headAnnotations;
      if (head) {
        const headRecord = await this.currentRevision(opened, documentId, head);
        const headMaterial = await this.materialFor(opened, headRecord);
        annotations = rebaseOpenAnnotations(
          headMaterial.document,
          document,
          headAnnotations,
        );
        document = { ...document, annotations };
      }
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
      throw new AgentDocxError(
        "REVISION_NOT_FOUND",
        "Draft guidance requires a revision",
      );
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
        blockId:
          snapshot.document.blocks[paragraph.index]?.id ??
          `b_${"0".repeat(36)}`,
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
    options: {
      renderer?: "deterministic" | "word" | "libreoffice" | "compare";
    } = {},
  ): Promise<PatchEvaluation> {
    const opened = await openStore(this.manifestPath);
    const record = await this.currentRevision(
      opened,
      patch.documentId,
      patch.baseRevision,
    );
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
      headMatchesBase:
        (await readHead(opened.storePath, patch.documentId)) === record.id,
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
      for (const edit of [...patch.edits].sort(
        (left, right) => right.start - left.start,
      ))
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
          const diagnostic =
            candidateMeasurement.deterministic.paragraphs?.find(
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
            pageCount:
              candidateMeasurement.pageCount - beforeMeasurement.pageCount,
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
                candidateValidation.summary.pass -
                beforeValidation.summary.pass,
              fail:
                candidateValidation.summary.fail -
                beforeValidation.summary.fail,
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
      throw new AgentDocxError(
        "PATCH_MISMATCH",
        "Patch hash does not match evaluation",
      );
    if (evaluation.candidate.status !== "ok" || !evaluation.canApply)
      throw new AgentDocxError(
        "PATCH_INVALID",
        "Patch cannot be applied to the working copy",
      );
    return withLockedStore(this.manifestPath, async (opened) => {
      const config = documentById(opened.manifest, patch.documentId);
      const record = await this.currentRevision(
        opened,
        patch.documentId,
        patch.baseRevision,
      );
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
      for (const edit of [...patch.edits].sort(
        (left, right) => right.start - left.start,
      ))
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
        candidateValidation.summary.unknown >
          beforeValidation.summary.unknown ||
        pageLimitExcess(candidateMeasurement) >
          pageLimitExcess(beforeMeasurement) ||
        countedLineExcess(candidateMeasurement) >
          countedLineExcess(beforeMeasurement);
      if (
        gate === "pass" &&
        (candidateValidation.status !== "pass" ||
          pageLimitExcess(candidateMeasurement) !== 0 ||
          countedLineExcess(candidateMeasurement) !== 0)
      )
        throw new AgentDocxError(
          "PATCH_FAILED_VALIDATION",
          "Patch does not pass constraints",
        );
      if (gate === "not-worse" && worse)
        throw new AgentDocxError(
          "PATCH_FAILED_VALIDATION",
          "Patch worsens the selected revision's constraints",
        );
      return this.commitLocked(
        opened,
        config,
        candidateSnapshot,
        document,
        annotations,
        record.id,
        input.author,
        input.message,
        undefined,
        true,
        true,
        { expectedWorkingTreeHash: snapshot.workingTreeHash },
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
      const base = await this.currentRevision(
        opened,
        documentId,
        input.baseRevision,
      );
      const target = await this.currentRevision(
        opened,
        documentId,
        input.targetRevision,
      );
      if ((await readHead(opened.storePath, documentId)) !== base.id)
        throw new AgentDocxError(
          "REVISION_CONFLICT",
          "Restore base is not the current head",
        );
      const currentConfig = documentById(opened.manifest, documentId);
      const currentSnapshot = await snapshotProjectDocument(
        opened,
        currentConfig,
      );
      if (currentSnapshot.workingTreeHash !== base.workingTreeHash)
        throw new AgentDocxError(
          "WORKING_COPY_CONFLICT",
          "Working copy differs from restore base",
        );
      const material = await this.materialFor(opened, target);
      if (material.config.source !== currentConfig.source)
        throw new AgentDocxError(
          "PROJECT_INVALID",
          "Restore cannot change the document source path",
        );
      const targetSnapshot = await snapshotWithDependencies(
        opened,
        currentSnapshot,
        material.config,
        target.dependencyObjects,
      );
      const snapshot = snapshotWithSource(targetSnapshot, material.source);
      return this.commitLocked(
        opened,
        material.config,
        snapshot,
        material.document,
        material.annotations,
        base.id,
        input.author,
        input.message,
        undefined,
        true,
        true,
        { expectedWorkingTreeHash: currentSnapshot.workingTreeHash },
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
        !(await this.isFirstParentAncestor(opened, base.id, head))
      )
        throw new AgentDocxError(
          "CHANGESET_INVALID",
          "Change-set base must be a distinct first-parent ancestor",
        );
      const baseMaterial = await this.materialFor(opened, base);
      const headMaterial = await this.materialFor(opened, head);
      const rawExpected = createChangeSet(
        documentId,
        base.id,
        head.id,
        baseMaterial.document,
        headMaterial.document,
        baseMaterial.annotations,
        headMaterial.annotations,
        defaultAttribution(head.author, head.createdAt),
        {
          baseConfig: baseMaterial.config as unknown as JsonObject,
          headConfig: headMaterial.config as unknown as JsonObject,
          baseDependencies: base.dependencyObjects,
          headDependencies: head.dependencyObjects,
          baseSource: baseMaterial.source,
          headSource: headMaterial.source,
        },
      );
      const baseProvenance = await provenanceForRevision(opened, base);
      const headProvenance = await provenanceForRevision(opened, head);
      const expected = reattributeChangeSet(rawExpected, {
        baseBlocks: baseProvenance.blocks,
        headBlocks: headProvenance.blocks,
        baseOperations: baseProvenance.operations,
        headOperations: headProvenance.operations,
        baseConfig: baseProvenance.config,
        headConfig: headProvenance.config,
        baseConfigOperations: baseProvenance.configOperations,
        headConfigOperations: headProvenance.configOperations,
        baseDependencyOperations: baseProvenance.dependencyOperations,
        headDependencyOperations: headProvenance.dependencyOperations,
        baseDocument: baseMaterial.document,
        headDocument: headMaterial.document,
        baseDependencies: baseProvenance.dependencies,
        headDependencies: headProvenance.dependencies,
      });
      if (canonicalJson(expected) !== canonicalJson(input.changeSet))
        throw new AgentDocxError(
          "CHANGESET_INVALID",
          "Change set does not match the selected immutable revisions",
        );
      if (
        base.id === head.id ||
        (await readHead(opened.storePath, documentId)) !== head.id
      )
        throw new AgentDocxError(
          "REVISION_CONFLICT",
          "Change-set head must be the distinct current document head",
        );
      const changeIds = [
        ...expected.changes.map((change) => change.id),
        ...expected.annotations.map((change) => change.id),
      ].sort();
      const decisionIds = Object.keys(input.decisions).sort();
      if (
        Object.values(input.decisions).some(
          (decision) => decision !== "accept" && decision !== "reject",
        )
      )
        throw new AgentDocxError(
          "CHANGESET_INVALID",
          "Change-set decisions must be accept or reject",
        );
      if (
        changeIds.length !== decisionIds.length ||
        changeIds.some((id, index) => id !== decisionIds[index])
      )
        throw new AgentDocxError(
          "CHANGESET_INVALID",
          "Change-set decisions must select every change exactly once",
        );
      const currentConfig = documentById(opened.manifest, documentId);
      const snapshot = await snapshotProjectDocument(opened, currentConfig);
      if (
        snapshot.workingTreeHash !== head.workingTreeHash ||
        snapshot.sourceObject !== head.sourceObject ||
        snapshot.documentConfigObject !== head.documentConfigObject ||
        canonicalJson(snapshot.dependencyObjects) !==
          canonicalJson(head.dependencyObjects)
      )
        throw new AgentDocxError(
          "WORKING_COPY_CONFLICT",
          "Working copy differs from the change-set head",
        );
      const targetConfig = applyRejectedConfigChanges(
        headMaterial.config,
        expected.changes,
        input.decisions,
      );
      const targetDependencies = applyRejectedDependencyChanges(
        head.dependencyObjects,
        expected.changes,
        input.decisions,
      );
      const replacements = rejectedSourceReplacements(
        snapshot.source,
        baseMaterial.source,
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
      const targetSnapshot = await snapshotWithDependencies(
        opened,
        snapshot,
        targetConfig,
        targetDependencies,
      );
      const preparedSnapshot = snapshotWithSource(targetSnapshot, source);
      const document = documentFor(
        source,
        targetConfig,
        preparedSnapshot,
        opened.manifest.projectId,
        annotations,
        true,
      );
      return this.commitLocked(
        opened,
        targetConfig,
        preparedSnapshot,
        document,
        annotations,
        head.id,
        input.author,
        input.message,
        { schemaVersion: 1, changeSet: expected, decisions: input.decisions },
        false,
        true,
        {
          expectedWorkingTreeHash: snapshot.workingTreeHash,
          parentIds: [base.id, head.id],
          firstParent: base,
        },
      );
    });
  }
  async exportDocx(
    documentId: string,
    input: ExportDocxInput,
  ): Promise<ProjectCompiledDocx> {
    return withLockedStore(this.manifestPath, async (opened) => {
      const record = await this.currentRevision(
        opened,
        documentId,
        input.revision,
      );
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
          {
            baseSource: baseMaterial.source,
            headSource: material.source,
          },
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
              pageCount: Math.max(
                1,
                compiled.measurement.deterministic.pageCount,
              ),
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

      const destination = await assertExportDestination(
        opened,
        material.config,
        input.output,
      );
      const owner = this.randomUuid();
      const stagePath = `${destination.output}.agent-docx-${owner}.stage`;
      if (await pathExists(stagePath))
        throw new AgentDocxError(
          "OUTPUT_EXISTS",
          `DOCX export stage already exists: ${stagePath}`,
        );
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
        artifactProvenanceSha256: emptyObject,
        artifactStorePath: resolve(
          opened.storePath,
          "artifacts",
          "0".repeat(64),
          "0".repeat(64),
          "document.docx",
        ),
        attachmentStorePath: null,
      };
      await updateExportIntent(opened.projectDirectory, initialIntent);
      let artifactObject!: RevisionId;
      let artifactProvenance!: RevisionId;
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
        if (compiled.attachments) {
          await writeObject(
            opened.storePath,
            canonicalJson(compiled.attachments.manifest),
          );
          for (const entry of compiled.attachments.manifest.entries)
            await writeObject(
              opened.storePath,
              compiled.attachments.files[entry.name]!.bytes,
            );
        }
        const provenanceJson = canonicalJson({
          schemaVersion: 1,
          generator: "agent-docx",
          generatorVersion: "0.1.0",
          documentId,
          revision: record.id,
          mode,
          baseRevision,
          profile: compiled.artifact.profile,
          rulePack: compiled.artifact.rulePack,
          dependencies: Object.fromEntries(
            Object.entries(record.dependencyObjects).sort(([left], [right]) =>
              left.localeCompare(right),
            ),
          ),
          docxSha256: artifactObject,
          attachments: compiled.attachments
            ? [...compiled.attachments.manifest.entries]
                .map((entry) => ({
                  name: entry.name,
                  mediaType: entry.mediaType,
                  byteLength: entry.byteLength,
                  sha256: entry.sha256,
                  payloadPath: entry.payloadPath,
                }))
                .sort((left, right) => left.name.localeCompare(right.name))
            : null,
          attachmentManifestSha256:
            compiled.attachments?.manifestSha256 ?? null,
        });
        artifactProvenance = objectId(provenanceJson);
        const artifactDirectory = artifactDirectoryFor(
          opened.storePath,
          record.id,
          artifactProvenance,
        );
        artifactStorePath = resolve(artifactDirectory, "document.docx");
        attachmentStorePath = compiled.attachments
          ? resolve(artifactDirectory, "attachments", "manifest.json")
          : null;
        const artifactStagePath = resolve(stagePath, "artifact");
        await mkdir(artifactStagePath, { mode: 0o700 });
        await writeExclusiveFile(
          resolve(artifactStagePath, "document.docx"),
          bytes,
        );
        await writeExclusiveFile(
          resolve(artifactStagePath, "provenance.json"),
          provenanceJson,
        );
        if (compiled.attachments)
          await writeAttachmentStage(artifactStagePath, compiled.attachments);
        await updateExportIntent(opened.projectDirectory, {
          ...initialIntent,
          state: "prepared",
          attachmentPath,
          attachmentStagePath: compiled.attachments
            ? resolve(stagePath, "attachments")
            : null,
          docxSha256: artifactObject,
          attachmentManifestSha256:
            compiled.attachments?.manifestSha256 ?? null,
          artifactProvenanceSha256: artifactProvenance,
          artifactStorePath,
          attachmentStorePath,
        });
        publicationStarted = true;
        await completeExportIntent(
          opened.projectDirectory,
          opened.manifestPath,
        );
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
        provenanceSha256: artifactProvenance,
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
      } as ProjectCompiledDocx;
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
        await assertNoSymlinkComponents(asset.path, "Imported asset");
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
            undefined,
            false,
            true,
            { expectedWorkingTreeHash: currentSnapshot.workingTreeHash },
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
            undefined,
            false,
            true,
            { expectedWorkingTreeHash: currentSnapshot.workingTreeHash },
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
        await replaceOwnedFile(
          sourcePath,
          snapshot.sourceObject,
          inspected.source,
        );
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
          undefined,
          false,
          true,
          { expectedWorkingTreeHash: refreshed.workingTreeHash },
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
          await replaceOwnedFile(
            sourcePath,
            refreshed.sourceObject,
            snapshot.source,
          );
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
  const originalSource = input.createSource
    ? null
    : await readFile(resolve(projectDirectory, config.source));
  const projectId = options.randomUUID?.() ?? systemRandomUuid();
  const manifest: AgentDocxManifest = {
    schemaVersion: 1,
    projectId,
    defaultDocument: config.id,
    storeDir: ".agent-docx",
    documents: [config],
  };
  await initializeStore(absoluteManifestPath, manifest);
  const sourcePath = resolve(projectDirectory, config.source);
  let createdSource = false;
  let sourceMaterialized = false;
  try {
    if (input.createSource) {
      await createEmptySource(projectDirectory, input.source);
      createdSource = true;
    }
    await withLockedStore(absoluteManifestPath, async (opened) => {
      await materializeSourceMarkers(opened, config);
      sourceMaterialized = true;
    });
  } catch (error) {
    try {
      if (originalSource === null && createdSource) {
        const entry = await lstat(sourcePath);
        if (entry.isFile() && !entry.isSymbolicLink())
          await removeOwnedFile(
            sourcePath,
            objectId(await readFile(sourcePath)),
          );
      } else if (originalSource !== null && sourceMaterialized) {
        const current = await readFile(sourcePath);
        await replaceOwnedFile(sourcePath, objectId(current), originalSource);
      }
      await removeInitializedProject(absoluteManifestPath, manifest);
    } catch (rollbackError) {
      throw rollbackError;
    }
    throw error;
  }
  return new Project(
    absoluteManifestPath,
    options.clock ?? (() => new Date()),
    options.randomUUID ?? systemRandomUuid,
  );
};
