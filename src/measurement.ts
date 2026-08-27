import type {
  Diagnostic,
  ErrorCode,
  JsonValue,
  SectionHeading,
  SourcePosition,
} from "./types.js";
import type {
  BuiltInProfileId,
  FontSetInput,
  LayoutOverrides,
  LayoutProfile,
  ResolvedLayoutProfile,
} from "./layout/profile.js";
import type { DocxTemplateInspection } from "./docx/contracts.js";
import type { DocumentChrome, LitigationMetadata } from "./legal/model.js";
import type { UserRulePack } from "./legal/rules.js";

export type FilingKind =
  | "principal-brief"
  | "reply-brief"
  | "motion-document"
  | "opposition-text"
  | "reply-text";

export type RendererMode = "deterministic" | "word" | "libreoffice" | "compare";
export type PageCountSource = "deterministic" | "word" | "libreoffice";
export type ProvenanceSource =
  | "rule"
  | "package"
  | "template"
  | "override"
  | "fallback";

export type EstimateOptions = {
  profile?: BuiltInProfileId | LayoutProfile;
  template?: DocxTemplateInspection;
  layout?: LayoutOverrides;
  fontSet?: FontSetInput;
  chrome?: DocumentChrome;
  metadata?: LitigationMetadata;
  assets?: Readonly<Record<string, { bytes: Uint8Array; mediaType: string }>>;
  filingKind?: FilingKind;
  pageLimit?: number;
  paragraphDiagnostics?: boolean;
  sectionDiagnostics?: boolean;
  trim?: false | { maxCandidates?: number; maxLastLineRatio?: number };
  lineDiagnostics?: boolean;
};

export type SectionPageDiagnostic = {
  page: number;
  bodyVisualLines: number;
  footnoteVisualLines: number;
  visualLines: number;
  countedLines: number;
};

export type SectionPageBudget = {
  limitPages: number;
  withinLimit: boolean;
  pagesBeyondLimit: readonly number[];
};

export type SectionDiagnostic = {
  source: "deterministic";
  index: number;
  parentIndex: number | null;
  heading: SectionHeading | null;
  position: SourcePosition | null;
  empty: boolean;
  startPage: number | null;
  endPage: number | null;
  pageCount: number;
  bodyVisualLines: number;
  footnoteVisualLines: number;
  visualLines: number;
  countedLines: number;
  pages: readonly SectionPageDiagnostic[];
  pageBudget?: SectionPageBudget;
};

export type WordRendererOptions = { powerShellPath?: string };
export type LibreOfficeRendererOptions = {
  executablePath?: string;
  installedFonts?: readonly { family: string; path: string }[];
  includePdfBytes?: boolean;
};

export type MeasureOptions = EstimateOptions & {
  renderer?: RendererMode;
  officeTimeoutMs?: number;
  word?: WordRendererOptions;
  libreoffice?: LibreOfficeRendererOptions;
  includeGeneratedDocx?: boolean;
  rulePacks?: readonly UserRulePack[];
};

export type LastPageMetrics = {
  source: "deterministic";
  visualLines: number;
  usedTwips: number;
  usableTwips: number;
  bodyLineEquivalentsUsed: number;
  bodyLineCapacity: number;
};

export type LineDiagnostic = {
  source: "deterministic";
  globalIndex: number;
  page: number;
  indexOnPage: number;
  blockIndex: number;
  position: SourcePosition;
  visualLinesInBlock: number;
  indexInBlock: number;
  isLastLineOfBlock: boolean;
  usedTwips: number;
  availableTwips: number;
  unusedTwips: number;
  ratio: number;
  overflowed: boolean;
  counted: boolean;
  text: string;
  start: number;
  end: number;
  contentEnd: number;
  startCause: "soft" | "hard" | "start";
};

export type ParagraphDiagnostic = {
  source: "deterministic";
  index: number;
  position: SourcePosition;
  startPage: number;
  endPage: number;
  visualLines: number;
  lastLineText: string;
  lastLineTextRange: { start: number; end: number };
  lastLineSourceRanges: readonly {
    position: SourcePosition;
    precision: "exact" | "node";
  }[];
  lastLineUsedTwips: number;
  lastLineAvailableTwips: number;
  lastLineUnusedTwips: number;
  lastLineRatio: number;
  lastLineOverflow: boolean;
  penultimateLineText: string | null;
  penultimateLineUnusedTwips: number | null;
  oneLineReduction: {
    estimatedRemovalTwips: number;
    basis: "deterministic-tail-width-deficit";
    confidence: "heuristic";
  } | null;
  preview: string;
};

export type TrimOpportunity = ParagraphDiagnostic & {
  rank: number;
  message: "This block may lose one wrapped line after removing or rephrasing approximately the reported width; verify by re-running pagination.";
};

export type Budget = {
  limitPages: number;
  withinLimit: boolean;
  pagesRemaining: number;
  equivalentPagesRemaining: number;
  bodyLineEquivalentsRemaining: number;
  fractionalFieldsSource: "deterministic";
};

export type DeterministicResult = {
  schemaVersion: 1;
  pageCount: number;
  equivalentPages: number;
  totalVisualLines: number;
  visualLinesByPage: readonly number[];
  countedLinesByPage: readonly number[];
  lastPage: LastPageMetrics | null;
  profile: ResolvedLayoutProfile;
  warnings: readonly Diagnostic[];
  paragraphs?: readonly ParagraphDiagnostic[];
  trimOpportunities?: readonly TrimOpportunity[];
  sections?: readonly SectionDiagnostic[];
  lines?: readonly LineDiagnostic[];
  budget?: Budget;
};

export type RendererError = {
  code: ErrorCode;
  message: string;
  phase: string;
  details?: Readonly<Record<string, JsonValue>>;
};

export type WordParagraphDiagnostic = {
  source: "word";
  index: number;
  position: SourcePosition;
  startPage: number;
  endPage: number;
  lineCount: number;
  finalLineText: string;
  preview: string;
};

export type RendererStatus<T> =
  | { status: "ok"; value: T }
  | { status: "unavailable" | "error"; error: RendererError };

export type WordRendering = {
  pageCount: number;
  totalBodyLines: number;
  bodyLinesByPage: readonly number[];
  bodyLinesOnLastPage: number | null;
  version: string;
  build: string;
  activePrinter: string;
  requestedFontFamilies: readonly string[];
  paragraphDiagnostics?:
    | { status: "ok"; value: readonly WordParagraphDiagnostic[] }
    | { status: "error"; error: RendererError };
  generatedDocxSha256: string;
  durationMs: number;
  cleanupState: "complete" | "forced" | "unverified";
};

export type LibreOfficeRendering = {
  pageCount: number;
  versionRaw: string;
  executablePath: string;
  platform: string;
  arch: string;
  calibratedFontEnvironment: boolean;
  requestedFontFamilies: readonly string[];
  generatedDocxSha256: string;
  pdfSha256: string;
  durationMs: number;
  pdf?: Uint8Array;
};

export type MeasurementResult = {
  schemaVersion: 1;
  mode: RendererMode;
  pageCount: number;
  pageCountSource: PageCountSource;
  deterministic: DeterministicResult;
  renderers: {
    word?: RendererStatus<WordRendering>;
    libreoffice?: RendererStatus<LibreOfficeRendering>;
  };
  budget?: Budget;
  generatedDocx?: Uint8Array;
  budgetBySource?: Readonly<Partial<Record<PageCountSource, Budget>>>;
};

/** Projection of a measurement result without the generated DOCX payload. */
export const serializableMeasurement = (
  measurement: MeasurementResult,
): Omit<MeasurementResult, "generatedDocx"> => {
  const { generatedDocx: _generatedDocx, ...serializable } = measurement;
  return serializable;
};
