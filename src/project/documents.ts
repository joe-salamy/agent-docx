import { AgentDocxError } from "../types.js";
import { toErrorPayload } from "../errors.js";
import type { PatchEvaluation, SourcePatch } from "../draft/types.js";
import type { UserRulePack } from "../legal/rules.js";
import type { ProjectContext } from "./context.js";
import { rm } from "node:fs/promises";
import { isAbsolute, relative, resolve, sep } from "node:path";
import { isSafeRelativePath } from "../path-util.js";
import {
  canonicalJson,
  canonicalObjectId,
  objectId,
  openStore,
  readHead,
  readProjectFile,
  readRevisionJson,
  replaceOwnedFile,
  snapshotProjectDocument,
  storeSnapshot,
  updateManifest,
  withLockedStore,
  writeHead,
  writeObject,
  writeRevisionJson,
  type OpenedStore,
  type ProjectSnapshot,
} from "./store.js";
import {
  annotationsForHead,
  currentRevision,
  dependencyPathsChanged,
  materialFor,
  isUtf16Boundary,
  materializeSelectedDependencies,
  provenanceForRevision,
  snapshotForMaterial,
} from "./revisions.js";
import {
  createRevisionDelta,
  defaultAttribution,
  rebaseOpenAnnotations,
  type JsonObject,
} from "../revisions/diff.js";
import { measureNormalizedDocument } from "../renderers/index.js";
import { definedProps } from "../json-contract.js";
import {
  insertMissingBlockMarkers,
  parseLegalMarkdown,
  type LegalAssetInput,
} from "../legal/parse.js";
import {
  validateLegalDocument,
  validateUserRulePack,
  type ValidationResult,
} from "../legal/rules.js";
import { lowerLegalDocument } from "../legal/lower.js";
import type {
  Actor,
  LegalDocument,
  ReviewAnnotation,
  RevisionId,
} from "../legal/model.js";
import type {
  ResolutionRecord,
  RevisionMutationResult,
  RevisionRecord,
} from "../revisions/types.js";
import {
  documentById,
  snapshotWithSource,
  sourcePathFor,
  version,
} from "./index.js";
import type {
  AgentDocxDocumentConfig,
  AgentDocxManifest,
  ConfigureDocumentInput,
  DocumentSnapshot,
  ProjectMeasureOptions,
  ProjectMeasurementResult,
  ProjectState,
} from "./contracts.js";
import type { DraftEvaluationError, DraftGuidance } from "../draft/types.js";
import {
  serializableMeasurement,
  type MeasurementResult,
} from "../measurement.js";

export type CommitOptions = {
  expectedWorkingTreeHash?: RevisionId;
  parentIds?: readonly RevisionId[];
  firstParent?: RevisionRecord;
};

export const documentFor = (
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
    ...(config.chrome !== undefined ? { chrome: config.chrome } : {}),
    assets: sourceAssets(snapshot),
    annotations,
    requireMarkers,
  }).document;

export const sourceAssets = (
  snapshot: ProjectSnapshot,
): Readonly<Record<string, LegalAssetInput>> => {
  const assets: Record<string, LegalAssetInput> = {};
  for (const [key, value] of snapshot.dependencyBytes) {
    if (!key.startsWith("asset/")) continue;
    assets[key.slice("asset/".length)] = value;
  }
  return assets;
};

export const configuredRulePacks = async (
  opened: OpenedStore,
  config: AgentDocxDocumentConfig,
  snapshot: ProjectSnapshot,
): Promise<readonly UserRulePack[]> => {
  const packs: UserRulePack[] = [];
  for (const [index, configuredPath] of (config.rulePacks ?? []).entries()) {
    const key = `rule-pack:${index}`;
    const expected = snapshot.dependencyObjects[key];
    if (!expected)
      throw new AgentDocxError(
        "PROJECT_INVALID",
        `Rule pack dependency is missing: ${configuredPath}`,
      );
    const path = resolve(opened.projectDirectory, configuredPath);
    const contained = relative(opened.projectDirectory, path);
    if (
      contained === "" ||
      contained === ".." ||
      contained.startsWith(`..${sep}`) ||
      isAbsolute(contained)
    )
      throw new AgentDocxError(
        "PATH_OUTSIDE_PROJECT",
        `Rule pack is outside the project: ${configuredPath}`,
      );
    let bytes: Uint8Array;
    try {
      bytes = await readProjectFile(
        path,
        `Rule pack ${index}`,
        opened.projectDirectory,
      );
    } catch (error) {
      if ((error as AgentDocxError).code === "INPUT_NOT_FOUND")
        throw new AgentDocxError(
          "PROJECT_INVALID",
          `Rule pack changed since snapshot: ${configuredPath}`,
        );
      throw error;
    }
    if (objectId(bytes) !== expected)
      throw new AgentDocxError(
        "PROJECT_INVALID",
        `Rule pack changed since snapshot: ${configuredPath}`,
      );
    let parsed: unknown;
    try {
      parsed = JSON.parse(
        new TextDecoder("utf-8", { fatal: true }).decode(bytes),
      );
    } catch {
      throw new AgentDocxError(
        "RULE_PACK_INVALID",
        `Rule pack is not valid JSON: ${configuredPath}`,
      );
    }
    packs.push(validateUserRulePack(parsed, `Rule pack ${configuredPath}`));
  }
  return packs;
};

export const sourceFontSet = (
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

export const materializeSourceMarkers = async (
  opened: OpenedStore,
  config: AgentDocxDocumentConfig,
): Promise<string> => {
  const snapshot = await snapshotProjectDocument(opened, config);
  const marked = insertMissingBlockMarkers(snapshot.source, {
    projectId: opened.manifest.projectId,
    documentId: config.id,
    metadata: config.metadata,
    ...(config.chrome !== undefined ? { chrome: config.chrome } : {}),
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

export const importedAssetDestinations = (
  opened: OpenedStore,
  config: AgentDocxDocumentConfig,
  assets: Readonly<Record<string, LegalAssetInput>>,
): readonly { source: string; path: string; bytes: Uint8Array }[] => {
  const entries = Object.entries(assets).sort(([left], [right]) =>
    left < right ? -1 : left > right ? 1 : 0,
  );
  if (entries.length === 0) return [];
  if (!config.assetsDir)
    throw new AgentDocxError(
      "DOCX_IMPORT_UNSUPPORTED",
      "Imported DOCX assets require the target document to configure assetsDir",
    );
  const root = resolve(opened.projectDirectory, config.assetsDir);
  return entries.map(([source, asset]) => {
    if (!isSafeRelativePath(source))
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

export const mutationError = (error: unknown): DraftEvaluationError => {
  const projected = toErrorPayload(error);
  if (error instanceof AgentDocxError) return projected;
  return { code: "PATCH_INVALID", message: projected.message };
};

export const assertStoredConfig = (
  value: unknown,
  path: string,
): AgentDocxDocumentConfig => {
  if (typeof value !== "object" || value === null || Array.isArray(value))
    return storedObjectCorrupt(path);
  const record = value as Record<string, unknown>;
  if (typeof record.id !== "string" || typeof record.source !== "string")
    return storedObjectCorrupt(path);
  return record as unknown as AgentDocxDocumentConfig;
};

export const assertStoredDocument = (
  value: unknown,
  path: string,
): LegalDocument => {
  if (typeof value !== "object" || value === null || Array.isArray(value))
    return storedObjectCorrupt(path);
  const record = value as Record<string, unknown>;
  if (
    record.schemaVersion !== 1 ||
    typeof record.projectId !== "string" ||
    typeof record.documentId !== "string"
  )
    return storedObjectCorrupt(path);
  return record as unknown as LegalDocument;
};

export const assertStoredAnnotations = (
  value: unknown,
  path: string,
): ReviewAnnotation[] => {
  if (
    !Array.isArray(value) ||
    value.some(
      (entry) =>
        typeof entry !== "object" ||
        entry === null ||
        typeof (entry as Record<string, unknown>).id !== "string",
    )
  )
    return storedObjectCorrupt(path);
  return value as ReviewAnnotation[];
};

export const storedObjectCorrupt = (path: string): never => {
  throw new AgentDocxError(
    "PROJECT_INVALID",
    `Stored object is corrupt: ${path}`,
  );
};

export const measureSnapshot = async (
  config: AgentDocxDocumentConfig,
  snapshot: ProjectSnapshot,
  document: LegalDocument,
  revision: RevisionId | null,
  options: ProjectMeasureOptions = {},
): Promise<ProjectMeasurementResult> => {
  const { includeGeneratedDocx, ...safeOptions } =
    options as ProjectMeasureOptions & {
      includeGeneratedDocx?: boolean;
    };
  void includeGeneratedDocx;
  const configFontSet = sourceFontSet(config, snapshot);
  const measurement = await measureNormalizedDocument(
    lowerLegalDocument(document),
    {
      ...definedProps(safeOptions),
      profile: config.profile,
      ...(config.filingKind !== undefined
        ? { filingKind: config.filingKind }
        : {}),
      ...(config.chrome !== undefined ? { chrome: config.chrome } : {}),
      ...(configFontSet !== undefined ? { fontSet: configFontSet } : {}),
      assets: sourceAssets(snapshot),
    },
  );
  return {
    ...measurement,
    documentId: config.id,
    revision,
    workingTreeHash: snapshot.workingTreeHash,
    dependencyObjects: snapshot.dependencyObjects,
  };
};

export const prepareWorkingDocument = async (
  opened: OpenedStore,
  config: AgentDocxDocumentConfig,
  snapshot: ProjectSnapshot,
  annotations: readonly ReviewAnnotation[],
  materializeMarkers: boolean,
): Promise<{
  snapshot: ProjectSnapshot;
  document: LegalDocument;
}> => {
  let source = snapshot.source;
  if (materializeMarkers) {
    const marked = insertMissingBlockMarkers(source, {
      projectId: opened.manifest.projectId,
      documentId: config.id,
      metadata: config.metadata,
      ...(config.chrome !== undefined ? { chrome: config.chrome } : {}),
      assets: sourceAssets(snapshot),
    });
    if (marked !== source) {
      snapshot = snapshotWithSource(snapshot, marked);
      source = marked;
    }
  }
  return {
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
};

export const commitLocked = async (
  ctx: ProjectContext,
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
): Promise<RevisionMutationResult> => {
  const head = await readHead(opened.storePath, config.id);
  if (base === null ? head !== null : head !== base)
    throw new AgentDocxError(
      "REVISION_CONFLICT",
      "Revision base does not match current head",
    );
  // Persist the snapshot objects before reading base material: content
  // addressing makes these writes idempotent, and a torn object from an
  // interrupted earlier write is replaced here so subsequent reads (which
  // verify hashes) succeed instead of poisoning the store.
  await storeSnapshot(opened, snapshot);
  const sourceObject = await writeObject(opened.storePath, snapshot.source);
  const headParent = head
    ? await readRevisionJson<RevisionRecord>(opened.storePath, head)
    : null;
  const parent = options.firstParent ?? headParent;
  const parentMaterial = parent ? await materialFor(opened, parent) : null;
  const headParentMaterial =
    headParent && headParent.id === parent?.id
      ? parentMaterial
      : headParent
        ? await materialFor(opened, headParent)
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
  const documentConfigObject = await writeObject(
    opened.storePath,
    canonicalJson(config),
  );
  const legalDocumentObject = await writeObject(
    opened.storePath,
    canonicalJson(committedDocument),
  );
  const annotationsObject = await writeObject(
    opened.storePath,
    canonicalJson(committedAnnotations),
  );
  const createdAt = ctx.clock().toISOString();
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
  const measurement = await measureSnapshot(
    config,
    snapshot,
    committedDocument,
    revisionId,
  );
  const customPacks = await configuredRulePacks(opened, config, snapshot);
  const validation = validateLegalDocument(committedDocument, {
    revision: revisionId,
    ...(config.rulePack !== undefined ? { rulePack: config.rulePack } : {}),
    customPacks,
    ...(config.filingKind !== undefined
      ? { filingKind: config.filingKind }
      : {}),
    measurement: serializableMeasurement(measurement),
  });
  const currentConfig = documentById(opened.manifest, config.id);
  const currentSnapshot = await snapshotProjectDocument(opened, currentConfig);
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
};

export const getStateLocked = async (
  ctx: ProjectContext,
  opened: OpenedStore,
): Promise<ProjectState> => {
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
            canonicalJson(record.dependencyObjects) ===
              canonicalJson(snapshot.dependencyObjects),
          all: record?.workingTreeHash === snapshot.workingTreeHash,
        },
      };
    }),
  );
  return {
    schemaVersion: 1,
    manifestPath: ctx.manifestPath,
    manifest: opened.manifest,
    documents,
    filingSets: opened.manifest.filingSets ?? [],
  };
};

export const getDocumentLocked = async (
  opened: OpenedStore,
  documentId: string,
  revision?: RevisionId | "HEAD",
): Promise<DocumentSnapshot> => {
  const config = documentById(opened.manifest, documentId);
  const head = await readHead(opened.storePath, documentId);
  if (revision !== undefined) {
    const record = await currentRevision(opened, documentId, revision);
    const material = await materialFor(opened, record);
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
  const working = await workingSnapshotLocked(opened, config, documentId);
  return {
    schemaVersion: 1,
    documentId,
    revision: null,
    head,
    source: working.snapshot.source,
    sourceSha256: working.snapshot.sourceObject,
    workingTreeHash: working.snapshot.workingTreeHash,
    documentConfig: config,
    dependencyObjects: working.snapshot.dependencyObjects,
    document: working.document,
    annotations: working.annotations,
  };
};

export const measureLocked = async (
  opened: OpenedStore,
  documentId: string,
  revision?: RevisionId | "HEAD",
  options: ProjectMeasureOptions = {},
): Promise<ProjectMeasurementResult> => {
  const config = documentById(opened.manifest, documentId);
  if (revision === undefined) {
    const working = await workingSnapshotLocked(opened, config, documentId);
    const customPacks = await configuredRulePacks(
      opened,
      config,
      working.snapshot,
    );
    return measureSnapshot(config, working.snapshot, working.document, null, {
      ...options,
      rulePacks: customPacks,
    });
  }
  const record = await currentRevision(opened, documentId, revision);
  const material = await materialFor(opened, record);
  const snapshot = await snapshotForMaterial(opened, material);
  const customPacks = await configuredRulePacks(
    opened,
    material.config,
    snapshot,
  );
  return measureSnapshot(
    material.config,
    snapshot,
    material.document,
    record.id,
    { ...options, rulePacks: customPacks },
  );
};

export const workingSnapshotLocked = async (
  opened: OpenedStore,
  config: AgentDocxDocumentConfig,
  documentId: string,
): Promise<{
  snapshot: ProjectSnapshot;
  document: LegalDocument;
  annotations: readonly ReviewAnnotation[];
}> => {
  const snapshot = await snapshotProjectDocument(opened, config);
  const headAnnotations = await annotationsForHead(opened, documentId);
  let document = documentFor(
    snapshot.source,
    config,
    snapshot,
    opened.manifest.projectId,
    [],
    false,
  );
  let annotations = headAnnotations;
  const head = await readHead(opened.storePath, documentId);
  if (head) {
    const headRecord = await currentRevision(opened, documentId, head);
    const headMaterial = await materialFor(opened, headRecord);
    annotations = rebaseOpenAnnotations(
      headMaterial.document,
      document,
      headAnnotations,
    );
    document = { ...document, annotations };
  }
  return { snapshot, document, annotations };
};

export const getDocument = async (
  ctx: ProjectContext,
  documentId: string,
  revision?: RevisionId | "HEAD",
): Promise<DocumentSnapshot> => {
  return withLockedStore(ctx.manifestPath, (opened) =>
    getDocumentLocked(opened, documentId, revision),
  );
};

export const configureDocument = async (
  ctx: ProjectContext,
  documentId: string,
  input: ConfigureDocumentInput,
): Promise<RevisionMutationResult> => {
  return withLockedStore(ctx.manifestPath, async (opened) => {
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
    if (input.changes.rulePacks === null) delete next.rulePacks;
    else if (input.changes.rulePacks !== undefined)
      next.rulePacks = input.changes.rulePacks;
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
    const annotations = await annotationsForHead(opened, documentId);
    const prepared = await prepareWorkingDocument(
      opened,
      next,
      snapshot,
      annotations,
      true,
    );
    return commitLocked(
      ctx,
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
};

export const checkpoint = async (
  ctx: ProjectContext,
  documentId: string,
  input: {
    baseRevision: RevisionId | "HEAD" | null;
    author: Actor;
    message: string;
  },
): Promise<RevisionMutationResult> => {
  return withLockedStore(ctx.manifestPath, async (opened) => {
    const config = documentById(opened.manifest, documentId);
    const head = await readHead(opened.storePath, documentId);
    const base = input.baseRevision === "HEAD" ? head : input.baseRevision;
    const snapshot = await snapshotProjectDocument(opened, config);
    const annotations = await annotationsForHead(opened, documentId);
    const prepared = await prepareWorkingDocument(
      opened,
      config,
      snapshot,
      annotations,
      true,
    );
    return commitLocked(
      ctx,
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
};

export const measure = async (
  ctx: ProjectContext,
  documentId: string,
  revision?: RevisionId | "HEAD",
  options: ProjectMeasureOptions = {},
): Promise<ProjectMeasurementResult> => {
  return withLockedStore(ctx.manifestPath, (opened) =>
    measureLocked(opened, documentId, revision, options),
  );
};

export const validateLocked = async (
  opened: OpenedStore,
  documentId: string,
  revision?: RevisionId | "HEAD",
): Promise<ValidationResult> => {
  const config = documentById(opened.manifest, documentId);
  let selectedConfig = config;
  let selectedSnapshot: ProjectSnapshot;
  let document: LegalDocument;
  let snapshotRevision: RevisionId | null;
  if (revision === undefined) {
    const working = await workingSnapshotLocked(opened, config, documentId);
    selectedSnapshot = working.snapshot;
    document = working.document;
    snapshotRevision = null;
  } else {
    const record = await currentRevision(opened, documentId, revision);
    const material = await materialFor(opened, record);
    selectedConfig = material.config;
    selectedSnapshot = await snapshotForMaterial(opened, material);
    document = material.document;
    snapshotRevision = record.id;
  }
  const measurement = await measureSnapshot(
    selectedConfig,
    selectedSnapshot,
    document,
    snapshotRevision,
    {},
  );
  if (revision === undefined && (selectedConfig.rulePacks?.length ?? 0) > 0) {
    const head = await readHead(opened.storePath, documentId);
    if (head) {
      const headRecord = await currentRevision(opened, documentId, head);
      for (const [index] of selectedConfig.rulePacks!.entries()) {
        const key = `rule-pack:${index}`;
        if (
          selectedSnapshot.dependencyObjects[key] !==
          headRecord.dependencyObjects[key]
        )
          throw new AgentDocxError(
            "PROJECT_INVALID",
            `Rule pack changed since snapshot: ${selectedConfig.rulePacks![index]}`,
          );
      }
    }
  }
  const customPacks = await configuredRulePacks(
    opened,
    selectedConfig,
    selectedSnapshot,
  );
  return validateLegalDocument(document, {
    revision: snapshotRevision,
    ...(selectedConfig.rulePack !== undefined
      ? { rulePack: selectedConfig.rulePack }
      : {}),
    customPacks,
    ...(selectedConfig.filingKind !== undefined
      ? { filingKind: selectedConfig.filingKind }
      : {}),
    measurement: serializableMeasurement(measurement),
  });
};

export const validate = async (
  ctx: ProjectContext,
  documentId: string,
  revision?: RevisionId | "HEAD",
): Promise<ValidationResult> => {
  return withLockedStore(ctx.manifestPath, (opened) =>
    validateLocked(opened, documentId, revision),
  );
};

export const getDraftGuidance = async (
  ctx: ProjectContext,
  documentId: string,
  revision?: RevisionId | "HEAD",
): Promise<DraftGuidance> => {
  const selected = revision ?? "HEAD";
  const snapshot = await getDocument(ctx, documentId, selected);
  if (!snapshot.revision)
    throw new AgentDocxError(
      "REVISION_NOT_FOUND",
      "Draft guidance requires a revision",
    );
  const measurement = await measure(ctx, documentId, snapshot.revision, {
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
        snapshot.document.blocks[paragraph.index]?.id ?? `b_${"0".repeat(36)}`,
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
};

export const evaluatePatch = async (
  ctx: ProjectContext,
  patch: SourcePatch,
  options: {
    renderer?: "deterministic" | "word" | "libreoffice" | "compare";
  } = {},
): Promise<PatchEvaluation> => {
  const opened = await openStore(ctx.manifestPath);
  const record = await currentRevision(
    opened,
    patch.documentId,
    patch.baseRevision,
  );
  const material = await materialFor(opened, record);
  const baseSnapshot = await snapshotForMaterial(opened, material);
  const config = documentById(opened.manifest, patch.documentId);
  const current = await snapshotProjectDocument(opened, config);
  const beforeMeasurement = await measureSnapshot(
    material.config,
    baseSnapshot,
    material.document,
    record.id,
    {
      ...(options.renderer !== undefined ? { renderer: options.renderer } : {}),
      paragraphDiagnostics: true,
      sectionDiagnostics: true,
    },
  );
  const baseCustomPacks = await configuredRulePacks(
    opened,
    material.config,
    baseSnapshot,
  );
  const beforeValidation = validateLegalDocument(material.document, {
    revision: record.id,
    ...(material.config.rulePack !== undefined
      ? { rulePack: material.config.rulePack }
      : {}),
    customPacks: baseCustomPacks,
    ...(material.config.filingKind !== undefined
      ? { filingKind: material.config.filingKind }
      : {}),
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
    const candidateFontSet = sourceFontSet(material.config, candidateSnapshot);
    const candidateMeasurement = await measureNormalizedDocument(
      lowerLegalDocument(candidateDocument),
      {
        ...(options.renderer !== undefined
          ? { renderer: options.renderer }
          : {}),
        profile: material.config.profile,
        ...(material.config.filingKind !== undefined
          ? { filingKind: material.config.filingKind }
          : {}),
        ...(candidateFontSet !== undefined
          ? { fontSet: candidateFontSet }
          : {}),
        paragraphDiagnostics: true,
        sectionDiagnostics: true,
      },
    );
    const candidateValidation = validateLegalDocument(candidateDocument, {
      revision: record.id,
      ...(material.config.rulePack !== undefined
        ? { rulePack: material.config.rulePack }
        : {}),
      customPacks: baseCustomPacks,
      ...(material.config.filingKind !== undefined
        ? { filingKind: material.config.filingKind }
        : {}),
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
};

export const applyPatch = async (
  ctx: ProjectContext,
  patch: SourcePatch,
  input: {
    patchHash: string;
    gate?: "report" | "not-worse" | "pass";
    author: Actor;
    message: string;
  },
): Promise<RevisionMutationResult> => {
  const evaluation = await evaluatePatch(ctx, patch);
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
  return withLockedStore(ctx.manifestPath, async (opened) => {
    const config = documentById(opened.manifest, patch.documentId);
    const record = await currentRevision(
      opened,
      patch.documentId,
      patch.baseRevision,
    );
    const material = await materialFor(opened, record);
    const baseSnapshot = await snapshotForMaterial(opened, material);
    const baseCustomPacks = await configuredRulePacks(
      opened,
      material.config,
      baseSnapshot,
    );
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
      measureSnapshot(
        material.config,
        baseSnapshot,
        material.document,
        record.id,
      ),
      measureSnapshot(config, candidateSnapshot, document, record.id),
    ]);
    const [beforeValidation, candidateValidation] = [
      validateLegalDocument(material.document, {
        revision: record.id,
        ...(material.config.rulePack !== undefined
          ? { rulePack: material.config.rulePack }
          : {}),
        customPacks: baseCustomPacks,
        ...(material.config.filingKind !== undefined
          ? { filingKind: material.config.filingKind }
          : {}),
        measurement: serializableMeasurement(beforeMeasurement),
      }),
      validateLegalDocument(document, {
        revision: record.id,
        ...(config.rulePack !== undefined ? { rulePack: config.rulePack } : {}),
        ...(config.filingKind !== undefined
          ? { filingKind: config.filingKind }
          : {}),
        customPacks: baseCustomPacks,
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
    return commitLocked(
      ctx,
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
};
