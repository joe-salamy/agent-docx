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
export {
  builtInRulePacks,
  validateLegalDocument,
  validateUserRulePack,
} from "./legal/rules.js";
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
  FilingSet,
  FilingSetSnapshot,
  FilingSetValidation,
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
  RedlineImportResult,
} from "./docx/contracts.js";
export type {
  BuiltInRulePack,
  RuleCheckKind,
  RuleCheckParams,
  UserRulePack,
  UserRulePackCheck,
  UserRulePackCountedLinesMaximum,
  UserRulePackLengthAlternative,
  UserRulePackLineSpacing,
  UserRulePackMarginMinimum,
  UserRulePackPageSize,
  UserRulePackReferenceIntegrity,
  UserRulePackRequiredBlock,
  UserRulePackRequiredFooter,
  UserRulePackRequiredMetadata,
  UserRulePackTypeface,
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
  CliErrorPayload,
  CliErrorRecord,
  CliFatalRecord,
  CliJsonlRequest,
  CliResultRecord,
  CliSource,
  CliTrigger,
  CliWatchEndRecord,
  CliWatchReadyRecord,
} from "./cli-contract.js";
export type {
  BuiltInProfileId,
  FontSetInput,
  LayoutOverrides,
  LayoutProfile,
  LineSpacing,
  MetricFont,
  PageGeometry,
  PaginationRules,
  ResolvedLayoutProfile,
  TableStyle,
  TextStyle,
  ThematicBreakStyle,
} from "./layout/profile.js";
export type {
  Budget,
  DeterministicResult,
  EstimateOptions,
  FilingKind,
  LastPageMetrics,
  LibreOfficeRendererOptions,
  LibreOfficeRendering,
  MeasureOptions,
  MeasurementResult,
  PageCountSource,
  ParagraphDiagnostic,
  ProvenanceSource,
  RendererError,
  RendererMode,
  RendererStatus,
  SectionDiagnostic,
  SectionPageBudget,
  SectionPageDiagnostic,
  TrimOpportunity,
  WordParagraphDiagnostic,
  WordRendererOptions,
  WordRendering,
} from "./measurement.js";
export type {
  Diagnostic,
  DiagnosticCode,
  ErrorCode,
  JsonValue,
  SectionHeading,
  SourcePosition,
} from "./types.js";
export type {
  DocxTemplateInspection,
  InspectTemplateOptions,
  InspectedSection,
  InspectedStyle,
} from "./docx/contracts.js";
