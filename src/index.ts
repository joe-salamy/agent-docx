export { estimateMarkdown } from "./estimate.js";
export { measureMarkdown } from "./renderers/index.js";
export { inspectDocxTemplate, DOCX_LIMITS } from "./docx/inspect.js";
export { builtInProfiles } from "./profiles.js";
export { AgentDocxError } from "./types.js";
export { compileMarkdown } from "./docx/compile.js";
export { generateDocx } from "./docx/generate.js";
export { generateRedlineDocx } from "./docx/redline.js";
export { inspectDocx } from "./docx/import.js";
export { createProject, openProject } from "./project/index.js";
export {
  agentActions,
  dispatchAgentRequest,
  parseAgentRequest,
  serializeAgentValue,
} from "./agent.js";
export {
  insertMissingBlockMarkers,
  parseLegalMarkdown,
} from "./legal/parse.js";
export { lowerLegalDocument } from "./legal/lower.js";
export { builtInRulePacks, validateLegalDocument } from "./legal/rules.js";
export { emptyValidationResult } from "./legal/rules.js";
export type {
  AddReviewInput,
  AgentDocxDocumentConfig,
  AgentDocxManifest,
  AgentDocxProject,
  CompileOptions,
  ConfigureDocumentInput,
  DependencyHashes,
  DocumentConfigUpdate,
  DocumentSnapshot,
  ProjectDocumentInput,
  ProjectFontSetConfig,
  ProjectMeasureOptions,
  ProjectMeasurementResult,
  ProjectState,
  ResolveChangesInput,
  ResolveReviewInput,
  SerializableMeasurementResult,
  SerializableProjectMeasurementResult,
} from "./project/contracts.js";
export type {
  Actor,
  AnnotationId,
  BlockId,
  DocumentChrome,
  LegalBlock,
  LegalDocument,
  LegalDocumentSpecification,
  LitigationMetadata,
  RevisionId,
  RulePackId,
  ReviewAnnotation,
} from "./legal/model.js";
export type {
  AnnotationChange,
  AttributionSpan,
  Change,
  ChangeAttribution,
  ChangeSet,
  ResolutionRecord,
  RevisionDeltaRecord,
  RevisionMutationResult,
  RevisionPage,
  RevisionRecord,
} from "./revisions/types.js";
export type {
  ArtifactAttachmentBundle,
  ArtifactResult,
  AttachmentInventoryEntry,
  AttachmentManifest,
  ProjectCompiledDocx,
  StatelessCompiledDocx,
  DocxImportResult,
  ExportDocxInput,
  ImportAttachmentBundle,
  ImportDocxInput,
  RendererProvenance,
  SerializableCompiledDocx,
} from "./docx/contracts.js";
export type {
  BuiltInRulePack,
  RuleCheckKind,
  ValidationFinding,
  ValidationResult,
} from "./legal/rules.js";
export type {
  DraftGuidance,
  PatchDeltas,
  PatchEvaluation,
  SourcePatch,
} from "./draft/types.js";
export type {
  AgentAction,
  AgentDispatchResult,
  AgentRequest,
} from "./agent.js";
export type {
  Budget,
  CliErrorPayload,
  CliErrorRecord,
  CliFatalRecord,
  CliJsonlRequest,
  CliResultRecord,
  CliSource,
  CliTrigger,
  CliWatchEndRecord,
  CliWatchReadyRecord,
  BuiltInProfileId,
  DeterministicResult,
  Diagnostic,
  DiagnosticCode,
  DocxTemplateInspection,
  ErrorCode,
  EstimateOptions,
  FilingKind,
  FontSetInput,
  InspectTemplateOptions,
  InspectedSection,
  InspectedStyle,
  JsonValue,
  LastPageMetrics,
  LayoutOverrides,
  LayoutProfile,
  LibreOfficeRendererOptions,
  LibreOfficeRendering,
  LineSpacing,
  MeasureOptions,
  MeasurementResult,
  MetricFont,
  PageCountSource,
  PageGeometry,
  PaginationRules,
  ParagraphDiagnostic,
  ProvenanceSource,
  RendererError,
  RendererMode,
  RendererStatus,
  ResolvedLayoutProfile,
  SectionDiagnostic,
  SectionHeading,
  SectionPageBudget,
  SectionPageDiagnostic,
  SourcePosition,
  TextStyle,
  TableStyle,
  ThematicBreakStyle,
  TrimOpportunity,
  WordRendererOptions,
  WordRendering,
  WordParagraphDiagnostic,
} from "./types.js";
