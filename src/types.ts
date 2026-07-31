export type JsonValue =
  | null
  | boolean
  | number
  | string
  | readonly JsonValue[]
  | { readonly [key: string]: JsonValue };
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
export type BuiltInProfileId =
  | "us-district-conventional"
  | "frap-32"
  | "cand-civil";
export type SourcePosition = {
  start: { line: number; column: number; offset: number };
  end: { line: number; column: number; offset: number };
};
export type LineSpacing =
  | { rule: "auto"; numerator: number; denominator: 240 }
  | { rule: "exact" | "atLeast"; twips: number };
export type TextStyle = {
  fontSizeTwips: number;
  bold: boolean;
  italic: boolean;
  lineSpacing: LineSpacing;
  beforeTwips: number;
  afterTwips: number;
  leftIndentTwips: number;
  rightIndentTwips: number;
  firstLineIndentTwips: number;
  hangingIndentTwips: number;
  keepWithNext: boolean;
  keepLines: boolean;
};
export type ThematicBreakStyle = {
  beforeTwips: number;
  afterTwips: number;
  thicknessTwips: number;
  keepWithNext: boolean;
};
export type TableStyle = {
  body: TextStyle;
  header: TextStyle;
  cellPaddingTwips: {
    top: number;
    right: number;
    bottom: number;
    left: number;
  };
  borderTwips: number;
  repeatHeader: boolean;
};
export type PageGeometry = {
  widthTwips: number;
  heightTwips: number;
  marginsTwips: { top: number; right: number; bottom: number; left: number };
  headerTwips: number;
  footerTwips: number;
  gutterTwips: number;
};
export type PaginationRules = {
  widowOrphanControl: boolean;
  widowLines: number;
  orphanLines: number;
  maxCountedLinesPerPage: number | null;
  lineCapExclusions: readonly ("footnote" | "blockquote")[];
};
export type LayoutOverrides = {
  page?: {
    widthTwips?: number;
    heightTwips?: number;
    marginsTwips?: Partial<PageGeometry["marginsTwips"]>;
    headerTwips?: number;
    footerTwips?: number;
    gutterTwips?: number;
  };
  requestedFontFamily?: string;
  body?: Partial<TextStyle>;
  headings?: Partial<
    Record<"1" | "2" | "3" | "4" | "5" | "6", Partial<TextStyle>>
  >;
  blockquote?: Partial<TextStyle>;
  list?: Partial<TextStyle>;
  footnote?: Partial<TextStyle>;
  pagination?: Partial<PaginationRules>;
  thematicBreak?: Partial<ThematicBreakStyle>;
  table?: {
    body?: Partial<TextStyle>;
    header?: Partial<TextStyle>;
    cellPaddingTwips?: Partial<TableStyle["cellPaddingTwips"]>;
    borderTwips?: number;
    repeatHeader?: boolean;
  };
};
export type DiagnosticCode =
  | "CONVENTIONAL_PROFILE_ONLY"
  | "PROFILE_CONSTRAINT_VIOLATION"
  | "FONT_STYLE_FACE_REUSED"
  | "MISSING_GLYPH"
  | "UNBREAKABLE_OVERFLOW"
  | "CAND_CPI_NOT_AUTOMATICALLY_VALIDATED"
  | "DOCX_MULTIPLE_SECTION_GEOMETRIES"
  | "DOCX_FONT_METRICS_UNAVAILABLE"
  | "DOCX_UNSUPPORTED_FEATURE"
  | "DOCX_IGNORED_UNSAFE_PART"
  | "DOCX_STYLE_CYCLE"
  | "DOCX_STYLE_PARENT_MISSING"
  | "FOOTNOTE_SPLIT_CONSTRAINT_RELAXED"
  | "TABLE_ROW_SPLIT_CONSTRAINT_RELAXED"
  | "TABLE_HEADER_REPEAT_CONSTRAINT_RELAXED"
  | "LIBREOFFICE_FONT_ENVIRONMENT_UNVERIFIED";
export type Diagnostic = {
  code: DiagnosticCode;
  severity: "warning";
  message: string;
  position?: SourcePosition;
  details?: Readonly<Record<string, JsonValue>>;
};
export type LayoutProfile = {
  id: string;
  label: string;
  effectiveDate: string | null;
  sourceUrl: string | null;
  sourceCitation: string;
  page: PageGeometry;
  requestedFontFamily: string;
  body: TextStyle;
  headings: Readonly<Record<"1" | "2" | "3" | "4" | "5" | "6", TextStyle>>;
  blockquote: TextStyle;
  list: TextStyle;
  footnote: TextStyle;
  pagination: PaginationRules;
  thematicBreak: ThematicBreakStyle;
  table: TableStyle;
  maxCharactersPerInch: number | null;
  filingPageLimits: Readonly<Partial<Record<FilingKind, number>>>;
  provenance: Readonly<
    Record<
      string,
      { source: ProvenanceSource; citation?: string; detail?: string }
    >
  >;
  warnings: readonly Diagnostic[];
};
export type InspectedSection = {
  index: number;
  page: PageGeometry;
  sourcePart: string;
};
export type InspectedStyle = {
  styleId: string | null;
  name: string | null;
  resolved: TextStyle;
  requestedFontFamily: string;
  provenance: Readonly<Record<string, ProvenanceSource | string>>;
};
export type DocxTemplateInspection = {
  imported: LayoutOverrides;
  sections: readonly InspectedSection[];
  selectedSection: number;
  styles: {
    body: InspectedStyle;
    headings: Readonly<
      Record<"1" | "2" | "3" | "4" | "5" | "6", InspectedStyle>
    >;
    quote: InspectedStyle;
    footnote: InspectedStyle;
    footnoteReference: InspectedStyle | null;
  };
  package: { sha256: string; mainPart: string; macroEnabled: boolean };
  warnings: readonly Diagnostic[];
};
export type FontSetInput = {
  family: string;
  regular: Uint8Array;
  bold?: Uint8Array;
  italic?: Uint8Array;
  boldItalic?: Uint8Array;
};
export type MetricFont = {
  role: "regular" | "bold" | "italic" | "boldItalic";
  requestedFamily: string;
  metricsFamily: string;
  sha256: string;
  substitutedMetrics: boolean;
};
export type ResolvedLayoutProfile = LayoutProfile & {
  metricFonts: readonly MetricFont[];
  template: null | {
    packageSha256: string;
    mainPart: string;
    selectedSection: number;
    macroEnabled: boolean;
    warnings: readonly Diagnostic[];
  };
};
export type EstimateOptions = {
  profile?: BuiltInProfileId | LayoutProfile;
  template?: DocxTemplateInspection;
  layout?: LayoutOverrides;
  fontSet?: FontSetInput;
  filingKind?: FilingKind;
  pageLimit?: number;
  paragraphDiagnostics?: boolean;
  sectionDiagnostics?: boolean;
  trim?: false | { maxCandidates?: number; maxLastLineRatio?: number };
};
export type SectionPageDiagnostic = {
  page: number;
  bodyVisualLines: number;
  footnoteVisualLines: number;
  visualLines: number;
  countedLines: number;
};
export type SectionHeading = {
  level: 1 | 2 | 3 | 4 | 5 | 6;
  title: string;
  position: SourcePosition;
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
};
export type MeasureOptions = EstimateOptions & {
  renderer?: RendererMode;
  officeTimeoutMs?: number;
  word?: WordRendererOptions;
  libreoffice?: LibreOfficeRendererOptions;
  includeGeneratedDocx?: boolean;
};
export type InspectTemplateOptions = {
  fallbackProfile?: BuiltInProfileId | LayoutProfile;
};
export type LastPageMetrics = {
  source: "deterministic";
  visualLines: number;
  usedTwips: number;
  usableTwips: number;
  bodyLineEquivalentsUsed: number;
  bodyLineCapacity: number;
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
  lastPage: LastPageMetrics | null;
  profile: ResolvedLayoutProfile;
  warnings: readonly Diagnostic[];
  paragraphs?: readonly ParagraphDiagnostic[];
  trimOpportunities?: readonly TrimOpportunity[];
  sections?: readonly SectionDiagnostic[];
  budget?: Budget;
};
export type ErrorCode =
  | "INVALID_ARGUMENT"
  | "INTERNAL_ERROR"
  | "INVALID_CONFIG"
  | "OUTPUT_EXISTS"
  | "OUTPUT_WRITE_FAILED"
  | "INPUT_NOT_FOUND"
  | "INPUT_NOT_UTF8"
  | "UNSUPPORTED_MARKDOWN"
  | "INVALID_LAYOUT"
  | "INVALID_FONT"
  | "DOCX_INVALID"
  | "DOCX_TOO_LARGE"
  | "DOCX_UNSAFE"
  | "DOCX_XML_LIMIT"
  | "WORD_NOT_FOUND"
  | "WORD_WSL_BRIDGE_UNAVAILABLE"
  | "WORD_RENDER_FAILED"
  | "WORD_TIMEOUT"
  | "WORD_CLEANUP_FAILED"
  | "LIBREOFFICE_NOT_FOUND"
  | "LIBREOFFICE_RENDER_FAILED"
  | "LIBREOFFICE_TIMEOUT"
  | "NO_OFFICE_RENDERER";
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
export type CliJsonlRequest =
  | { id?: string | number | null; path: string }
  | { id?: string | number | null; name?: string; markdown: string };
export type CliSource =
  | { kind: "file"; path: string }
  | { kind: "stdin" }
  | { kind: "inline"; name: string | null };
export type CliTrigger = {
  kind: "initial" | "source-change" | "dependency-change";
  paths: readonly string[];
};
export type CliErrorPayload = {
  code: ErrorCode;
  message: string;
  details?: Record<string, JsonValue>;
};
export type CliResultRecord = {
  schemaVersion: 1;
  kind: "result";
  mode: "batch" | "watch";
  sequence: number;
  requestId: string | number | null;
  source: CliSource;
  trigger: CliTrigger | null;
  measurement: MeasurementResult;
};
export type CliErrorRecord = {
  schemaVersion: 1;
  kind: "error";
  mode: "batch" | "watch";
  sequence: number;
  requestId: string | number | null;
  source: CliSource;
  trigger: CliTrigger | null;
  error: CliErrorPayload;
};
export type CliWatchReadyRecord = {
  schemaVersion: 1;
  kind: "ready";
  mode: "watch";
  sequence: number;
  source: CliSource;
  dependencies: readonly string[];
};
export type CliWatchEndRecord = {
  schemaVersion: 1;
  kind: "end";
  mode: "watch";
  sequence: number;
  source: CliSource;
  reason: "SIGINT" | "SIGTERM";
};
export type CliFatalRecord = {
  schemaVersion: 1;
  kind: "fatal";
  error: CliErrorPayload;
};
export class AgentDocxError extends Error {
  readonly code: ErrorCode;
  readonly details?: Record<string, JsonValue>;
  constructor(
    code: ErrorCode,
    message: string,
    details?: Record<string, JsonValue>,
  ) {
    super(message);
    this.name = "AgentDocxError";
    this.code = code;
    if (details !== undefined) this.details = details;
  }
}
