import type {
  BuiltInProfileId,
  FontSetInput,
  LayoutProfile,
} from "../layout/profile.js";
import type {
  FilingKind,
  MeasureOptions,
  MeasurementResult,
} from "../measurement.js";
import type {
  Actor,
  AnnotationId,
  BlockId,
  DocumentChrome,
  LegalDocument,
  ReviewAnnotation,
  LitigationMetadata,
  RevisionId,
  RulePackId,
} from "../legal/model.js";
import type { ValidationResult } from "../legal/rules.js";
import type {
  ChangeSet,
  RevisionMutationResult,
  RevisionPage,
  RevisionRecord,
} from "../revisions/types.js";
import type {
  DraftGuidance,
  PatchEvaluation,
  SourcePatch,
} from "../draft/types.js";
import type {
  DocxImportResult,
  ExportDocxInput,
  ImportAttachmentBundle,
  ImportDocxInput,
  ProjectCompiledDocx,
  RedlineImportResult,
} from "../docx/contracts.js";

export type ProjectFontSetConfig = {
  family: string;
  regularPath: string;
  boldPath?: string;
  italicPath?: string;
  boldItalicPath?: string;
};

export type AgentDocxDocumentConfig = {
  id: string;
  source: string;
  profile: BuiltInProfileId;
  filingKind?: FilingKind;
  rulePack?: RulePackId;
  rulePacks?: readonly string[];
  template?: string;
  assetsDir?: string;
  fontSet?: ProjectFontSetConfig;
  chrome?: DocumentChrome;
  metadata: LitigationMetadata;
};

export type AgentDocxManifest = {
  schemaVersion: 1;
  projectId: string;
  defaultDocument: string;
  storeDir: ".agent-docx";
  documents: AgentDocxDocumentConfig[];
  filingSets?: readonly FilingSet[];
};

export type ProjectDocumentInput = {
  documentId: string;
  source: string;
  template?: string;
  assetsDir?: string;
  fontSet?: ProjectFontSetConfig;
  createSource?: boolean;
  profile: BuiltInProfileId;
  filingKind?: FilingKind;
  rulePack?: RulePackId;
  rulePacks?: readonly string[];
  metadata: LitigationMetadata;
  chrome?: DocumentChrome;
};

export type DocumentConfigUpdate = {
  profile?: BuiltInProfileId;
  filingKind?: FilingKind | null;
  rulePack?: RulePackId | null;
  rulePacks?: readonly string[] | null;
  template?: string | null;
  assetsDir?: string | null;
  fontSet?: ProjectFontSetConfig | null;
  metadata?: LitigationMetadata;
  chrome?: DocumentChrome | null;
};

export type ConfigureDocumentInput = {
  baseRevision: RevisionId | "HEAD" | null;
  changes: DocumentConfigUpdate;
  author: Actor;
  message: string;
};

export type DependencyHashes = Readonly<Record<string, RevisionId>>;
export type SerializableMeasurementResult = Omit<
  MeasurementResult,
  "generatedDocx"
>;

export type ProjectMeasurementResult = MeasurementResult & {
  documentId: string;
  revision: RevisionId | null;
  workingTreeHash: RevisionId;
  dependencyObjects: DependencyHashes;
};

export type SerializableProjectMeasurementResult = Omit<
  ProjectMeasurementResult,
  "generatedDocx"
>;

export type ProjectMeasureOptions = Omit<
  MeasureOptions,
  | "profile"
  | "filingKind"
  | "template"
  | "layout"
  | "fontSet"
  | "pageLimit"
  | "includeGeneratedDocx"
>;

export type CompileOptions = Omit<
  MeasureOptions,
  | "profile"
  | "filingKind"
  | "template"
  | "layout"
  | "fontSet"
  | "includeGeneratedDocx"
>;

export type ProjectState = {
  schemaVersion: 1;
  manifestPath: string;
  manifest: AgentDocxManifest;
  documents: readonly {
    documentId: string;
    head: RevisionId | null;
    headWorkingTreeHash: RevisionId | null;
    workingTreeHash: RevisionId;
    sourceSha256: RevisionId;
    documentConfigSha256: RevisionId;
    dependencyObjects: DependencyHashes;
    matchesHead: {
      source: boolean;
      documentConfig: boolean;
      dependencies: boolean;
      all: boolean;
    };
  }[];
  filingSets: readonly FilingSet[];
};

export type DocumentSnapshot = {
  schemaVersion: 1;
  documentId: string;
  revision: RevisionId | null;
  head: RevisionId | null;
  source: string;
  sourceSha256: RevisionId;
  workingTreeHash: RevisionId;
  documentConfig: AgentDocxDocumentConfig;
  dependencyObjects: DependencyHashes;
  document: LegalDocument;
  annotations: readonly ReviewAnnotation[];
};

export type ResolveChangesInput = {
  changeSet: ChangeSet;
  decisions: Readonly<Record<`c_${string}`, "accept" | "reject">>;
  author: Actor;
  message: string;
};

export type AddReviewInput = {
  revision: RevisionId | "HEAD";
  blockId: BlockId;
  range?: { start: number; length: number };
  author: Actor;
  message: string;
};

export type ResolveReviewInput = {
  revision: RevisionId | "HEAD";
  annotationId: AnnotationId;
  author: Actor;
  message: string;
};

export type AgentDocxProject = {
  getState(): Promise<ProjectState>;
  addDocument(
    input: ProjectDocumentInput & { makeDefault?: boolean },
  ): Promise<ProjectState>;
  configureDocument(
    documentId: string,
    input: ConfigureDocumentInput,
  ): Promise<RevisionMutationResult>;
  getDocument(
    documentId: string,
    revision?: RevisionId | "HEAD",
  ): Promise<DocumentSnapshot>;
  getDraftGuidance(
    documentId: string,
    revision?: RevisionId | "HEAD",
  ): Promise<DraftGuidance>;
  checkpoint(
    documentId: string,
    input: {
      baseRevision: RevisionId | "HEAD" | null;
      author: Actor;
      message: string;
    },
  ): Promise<RevisionMutationResult>;
  listRevisions(
    documentId: string,
    input?: { limit?: number; cursor?: RevisionId },
  ): Promise<RevisionPage>;
  getRevision(
    documentId: string,
    revision: RevisionId | "HEAD",
  ): Promise<RevisionRecord>;
  restore(
    documentId: string,
    input: {
      baseRevision: RevisionId | "HEAD";
      targetRevision: RevisionId | "HEAD";
      author: Actor;
      message: string;
    },
  ): Promise<RevisionMutationResult>;
  diff(
    documentId: string,
    base: RevisionId | "HEAD",
    head: RevisionId | "HEAD",
  ): Promise<ChangeSet>;
  resolveChanges(
    documentId: string,
    input: ResolveChangesInput,
  ): Promise<RevisionMutationResult>;
  evaluatePatch(
    patch: SourcePatch,
    options?: {
      renderer?: "deterministic" | "word" | "libreoffice" | "compare";
    },
  ): Promise<PatchEvaluation>;
  applyPatch(
    patch: SourcePatch,
    input: {
      patchHash: string;
      gate?: "report" | "not-worse" | "pass";
      author: Actor;
      message: string;
    },
  ): Promise<RevisionMutationResult>;
  addReview(
    documentId: string,
    input: AddReviewInput,
  ): Promise<RevisionMutationResult>;
  resolveReview(
    documentId: string,
    input: ResolveReviewInput,
  ): Promise<RevisionMutationResult>;
  measure(
    documentId: string,
    revision?: RevisionId | "HEAD",
    options?: ProjectMeasureOptions,
  ): Promise<ProjectMeasurementResult>;
  validate(
    documentId: string,
    revision?: RevisionId | "HEAD",
  ): Promise<ValidationResult>;
  exportDocx(
    documentId: string,
    input: ExportDocxInput,
  ): Promise<ProjectCompiledDocx>;
  importDocx(
    input: Extract<ImportDocxInput, { inspectOnly: false }>,
  ): Promise<DocxImportResult>;
  importRedline(input: {
    documentId: string;
    input: string | Uint8Array;
    attachments?: ImportAttachmentBundle;
    author: Actor;
    message: string;
  }): Promise<RedlineImportResult>;
  addFilingSet(input: {
    id: string;
    label?: string;
    documentIds: readonly string[];
    pageCap?: number;
  }): Promise<ProjectState>;
  removeFilingSet(id: string): Promise<ProjectState>;
  getFilingSet(id: string): Promise<FilingSetSnapshot>;
  validateFilingSet(id: string): Promise<FilingSetValidation>;
};

export type ResolveProfileInput = BuiltInProfileId | LayoutProfile;
export type ResolvedProjectFontSet = FontSetInput;

export type FilingSet = {
  id: string;
  label?: string;
  documentIds: readonly string[];
  pageCap?: number;
};

export type FilingSetSnapshot = {
  schemaVersion: 1;
  id: string;
  label: string | null;
  documentIds: readonly string[];
  pageCap: number | null;
  documents: readonly {
    documentId: string;
    head: RevisionId | null;
    workingTreeHash: RevisionId;
    matchesHead: boolean;
  }[];
};

export type FilingSetValidation = {
  schemaVersion: 1;
  id: string;
  documents: readonly {
    documentId: string;
    head: RevisionId | null;
    validation: ValidationResult | null;
    pageCount: number | null;
  }[];
  pageCap: {
    limit: number;
    totalPages: number;
    status: "pass" | "fail" | "unknown";
    detail: string;
  } | null;
  status: "pass" | "fail" | "unknown";
};
