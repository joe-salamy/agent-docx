import { randomUUID as systemRandomUuid } from "node:crypto";
import { lstat } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import {
  currentRevision,
  diff,
  isUtf16Boundary,
  getRevision,
  listRevisions,
  materialFor,
  resolveChanges,
  restore,
} from "./revisions.js";
import {
  applyPatch,
  checkpoint,
  commitLocked,
  configureDocument,
  evaluatePatch,
  getDocument,
  getDraftGuidance,
  getStateLocked,
  materializeSourceMarkers,
  measure,
  validate,
} from "./documents.js";
import { exportDocx } from "./export.js";
import { importDocx, importRedline } from "./import.js";
import {
  addFilingSet,
  getFilingSet,
  removeFilingSet,
  validateFilingSet,
} from "./filing-sets.js";
import type { ProjectContext } from "./context.js";
import { visibleTextForBlock } from "../legal/visible-text.js";
import type { ValidationResult } from "../legal/rules.js";
import {
  type Actor,
  type RevisionId,
  type ReviewAnnotation,
} from "../legal/model.js";
import { AgentDocxError } from "../types.js";
import type {
  ChangeSet,
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
  FilingSetSnapshot,
  FilingSetValidation,
  ProjectDocumentInput,
  ProjectMeasureOptions,
  ProjectMeasurementResult,
  ProjectState,
  ResolveChangesInput,
  ResolveReviewInput,
} from "./contracts.js";
import {
  createEmptySource,
  canonicalObjectId,
  documentConfigFromInput,
  initializeStore,
  objectId,
  openStore,
  readHead,
  readObject,
  readProjectFile,
  removeInitializedProject,
  removeOwnedFile,
  replaceOwnedFile,
  snapshotProjectDocument,
  updateManifest,
  withLockedStore,
  type OpenedStore,
  type ProjectSnapshot,
} from "./store.js";
import type {
  ProjectCompiledDocx,
  ExportDocxInput,
  ImportDocxInput,
  DocxImportResult,
  ImportAttachmentBundle,
  RedlineImportResult,
} from "../docx/contracts.js";
import type {
  DraftGuidance,
  PatchEvaluation,
  SourcePatch,
} from "../draft/types.js";

export type ProjectRuntimeOptions = {
  clock?: () => Date;
  randomUUID?: () => string;
};

export const version = "0.1.1";

export const sourcePathFor = (
  opened: OpenedStore,
  config: AgentDocxDocumentConfig,
): string => resolve(opened.projectDirectory, config.source);
export const snapshotWithSource = (
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

export const snapshotWithDependencies = async (
  opened: OpenedStore,
  snapshot: ProjectSnapshot,
  config: AgentDocxDocumentConfig,
  dependencyObjects: Readonly<Record<string, RevisionId>>,
): Promise<ProjectSnapshot> => {
  const sortedDependencies = Object.fromEntries(
    Object.entries(dependencyObjects).sort(([left], [right]) =>
      left < right ? -1 : left > right ? 1 : 0,
    ),
  ) as Readonly<Record<string, RevisionId>>;
  const dependencyBytes = new Map<
    string,
    { bytes: Uint8Array; mediaType: string }
  >();
  for (const [key, dependencyObject] of Object.entries(sortedDependencies))
    dependencyBytes.set(key, {
      bytes: await readObject(opened.storePath, dependencyObject),
      mediaType:
        snapshot.dependencyBytes.get(key)?.mediaType ?? storedMediaType(key),
    });
  const documentConfigObject = canonicalObjectId(config);
  return {
    ...snapshot,
    documentConfigObject,
    dependencyObjects: sortedDependencies,
    dependencyBytes,
    workingTreeHash: canonicalObjectId({
      sourceObject: snapshot.sourceObject,
      documentConfigObject,
      dependencyObjects: sortedDependencies,
    }),
  };
};
export const documentById = (
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
export const storedMediaType = (key: string): string => {
  if (key === "template")
    return "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
  if (key.startsWith("font/")) return "font/ttf";
  const normalized = key.toLowerCase();
  if (normalized.endsWith(".png")) return "image/png";
  if (normalized.endsWith(".jpg") || normalized.endsWith(".jpeg"))
    return "image/jpeg";
  if (normalized.endsWith(".pdf")) return "application/pdf";
  if (key.startsWith("rule-pack:")) return "application/json";
  return key.startsWith("rule-source/")
    ? "text/plain"
    : "application/octet-stream";
};

class Project implements AgentDocxProject {
  private readonly ctx: ProjectContext;
  constructor(
    manifestPath: string,
    clock: () => Date,
    randomUuid: () => string,
  ) {
    this.ctx = { manifestPath, clock, randomUuid };
  }

  async getState(): Promise<ProjectState> {
    return withLockedStore(this.ctx.manifestPath, (opened) =>
      getStateLocked(this.ctx, opened),
    );
  }

  async addDocument(
    input: ProjectDocumentInput & { makeDefault?: boolean },
  ): Promise<ProjectState> {
    return withLockedStore(this.ctx.manifestPath, async (opened) => {
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
        : await readProjectFile(
            sourcePath,
            "Document source",
            opened.projectDirectory,
          );
      if (input.createSource)
        await createEmptySource(opened.projectDirectory, input.source);
      const manifest: AgentDocxManifest = {
        ...opened.manifest,
        defaultDocument: input.makeDefault
          ? config.id
          : opened.manifest.defaultDocument,
        documents: [...opened.manifest.documents, config],
      };
      let next: OpenedStore;
      try {
        await materializeSourceMarkers(opened, config);
        next = await updateManifest(opened, manifest);
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
      return getStateLocked(this.ctx, next);
    });
  }

  async addFilingSet(input: {
    id: string;
    label?: string;
    documentIds: readonly string[];
    pageCap?: number;
  }): Promise<ProjectState> {
    return addFilingSet(this.ctx, input);
  }
  async removeFilingSet(id: string): Promise<ProjectState> {
    return removeFilingSet(this.ctx, id);
  }
  async getFilingSet(id: string): Promise<FilingSetSnapshot> {
    return getFilingSet(this.ctx, id);
  }
  async validateFilingSet(id: string): Promise<FilingSetValidation> {
    return validateFilingSet(this.ctx, id);
  }
  async configureDocument(
    documentId: string,
    input: ConfigureDocumentInput,
  ): Promise<RevisionMutationResult> {
    return configureDocument(this.ctx, documentId, input);
  }
  async getDocument(
    documentId: string,
    revision?: RevisionId | "HEAD",
  ): Promise<DocumentSnapshot> {
    return getDocument(this.ctx, documentId, revision);
  }
  async checkpoint(
    documentId: string,
    input: {
      baseRevision: RevisionId | "HEAD" | null;
      author: Actor;
      message: string;
    },
  ): Promise<RevisionMutationResult> {
    return checkpoint(this.ctx, documentId, input);
  }
  async listRevisions(
    documentId: string,
    input: { limit?: number; cursor?: RevisionId } = {},
  ): Promise<RevisionPage> {
    return listRevisions(this.ctx, documentId, input);
  }
  async getRevision(
    documentId: string,
    revision: RevisionId | "HEAD",
  ): Promise<RevisionRecord> {
    return getRevision(this.ctx, documentId, revision);
  }
  async diff(
    documentId: string,
    base: RevisionId | "HEAD",
    head: RevisionId | "HEAD",
  ): Promise<ChangeSet> {
    return diff(this.ctx, documentId, base, head);
  }
  async addReview(
    documentId: string,
    input: AddReviewInput,
  ): Promise<RevisionMutationResult> {
    return withLockedStore(this.ctx.manifestPath, async (opened) => {
      const record = await currentRevision(opened, documentId, input.revision);
      const head = await readHead(opened.storePath, documentId);
      if (head !== record.id)
        throw new AgentDocxError(
          "REVISION_CONFLICT",
          "Review must target the current head",
        );
      // The public input contract is {start, length}; end is derived here so
      // an inverted range is structurally inexpressible.
      const range =
        input.range === undefined
          ? undefined
          : {
              start: input.range.start,
              end: input.range.start + input.range.length,
            };
      const material = await materialFor(opened, record);
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
        range &&
        (!isUtf16Boundary(visibleTextForBlock(block), range.start) ||
          !isUtf16Boundary(visibleTextForBlock(block), range.end) ||
          range.start > range.end)
      )
        throw new AgentDocxError(
          "ANNOTATION_CONFLICT",
          "Review range must be a code-point-safe range within its block",
        );
      if (range && range.start >= range.end)
        throw new AgentDocxError(
          "ANNOTATION_CONFLICT",
          "Review text range must select at least one code point",
        );
      if (
        range &&
        material.annotations.some(
          (annotation) =>
            annotation.status === "open" &&
            annotation.blockId === input.blockId &&
            annotation.range !== undefined &&
            range!.start < annotation.range.end &&
            annotation.range.start < range!.end,
        )
      )
        throw new AgentDocxError(
          "ANNOTATION_CONFLICT",
          "Open review text ranges must not overlap within a block",
        );
      if (
        range &&
        !["paragraph", "blockquote", "heading", "numbered-paragraph"].includes(
          block.kind,
        ) &&
        !(block.kind === "footnote" && block.paragraphs.length === 1)
      )
        throw new AgentDocxError(
          "ANNOTATION_CONFLICT",
          "Review ranges must stay within one source-mapped paragraph",
        );
      const annotationId = `a_${this.ctx.randomUuid()}` as `a_${string}`;
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
        ...(range ? { range } : {}),
        author: input.author,
        createdAt: this.ctx.clock().toISOString(),
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
      return commitLocked(
        this.ctx,
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
    return withLockedStore(this.ctx.manifestPath, async (opened) => {
      const record = await currentRevision(opened, documentId, input.revision);
      const head = await readHead(opened.storePath, documentId);
      if (head !== record.id)
        throw new AgentDocxError(
          "REVISION_CONFLICT",
          "Review must target the current head",
        );
      const material = await materialFor(opened, record);
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
      return commitLocked(
        this.ctx,
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
    options?: ProjectMeasureOptions,
  ): Promise<ProjectMeasurementResult> {
    return measure(this.ctx, documentId, revision, options);
  }
  async validate(
    documentId: string,
    revision?: RevisionId | "HEAD",
  ): Promise<ValidationResult> {
    return validate(this.ctx, documentId, revision);
  }
  async getDraftGuidance(
    documentId: string,
    revision?: RevisionId | "HEAD",
  ): Promise<DraftGuidance> {
    return getDraftGuidance(this.ctx, documentId, revision);
  }
  async evaluatePatch(
    patch: SourcePatch,
    options?: {
      renderer?: "deterministic" | "word" | "libreoffice" | "compare";
    },
  ): Promise<PatchEvaluation> {
    return evaluatePatch(this.ctx, patch, options);
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
    return applyPatch(this.ctx, patch, input);
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
    return restore(this.ctx, documentId, input);
  }
  async resolveChanges(
    documentId: string,
    input: ResolveChangesInput,
  ): Promise<RevisionMutationResult> {
    return resolveChanges(this.ctx, documentId, input);
  }
  async exportDocx(
    documentId: string,
    input: ExportDocxInput,
  ): Promise<ProjectCompiledDocx> {
    return exportDocx(this.ctx, documentId, input);
  }
  async importDocx(
    input: Extract<ImportDocxInput, { inspectOnly: false }>,
  ): Promise<DocxImportResult> {
    return importDocx(this.ctx, input);
  }
  async importRedline(input: {
    documentId: string;
    input: string | Uint8Array;
    attachments?: ImportAttachmentBundle;
    author: Actor;
    message: string;
  }): Promise<RedlineImportResult> {
    return importRedline(this.ctx, input);
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
    : await readProjectFile(
        resolve(projectDirectory, config.source),
        "Document source",
        projectDirectory,
      );
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
            objectId(
              await readProjectFile(
                sourcePath,
                "Document source",
                projectDirectory,
              ),
            ),
          );
      } else if (originalSource !== null && sourceMaterialized) {
        const current = await readProjectFile(
          sourcePath,
          "Document source",
          projectDirectory,
        );
        await replaceOwnedFile(sourcePath, objectId(current), originalSource);
      }
      await removeInitializedProject(absoluteManifestPath, manifest);
    } catch (rollbackError) {
      throw new AgentDocxError(
        "INTERNAL_ERROR",
        `createProject failed: ${error instanceof Error ? error.message : String(error)}; rollback also failed: ${rollbackError instanceof Error ? rollbackError.message : String(rollbackError)}`,
      );
    }
    throw error;
  }
  return new Project(
    absoluteManifestPath,
    options.clock ?? (() => new Date()),
    options.randomUUID ?? systemRandomUuid,
  );
};
