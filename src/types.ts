export type JsonValue =
  | null
  | boolean
  | number
  | string
  | readonly JsonValue[]
  | { readonly [key: string]: JsonValue };

export type SourcePosition = {
  start: { line: number; column: number; offset: number };
  end: { line: number; column: number; offset: number };
};

export type SectionHeading = {
  level: 1 | 2 | 3 | 4 | 5 | 6;
  title: string;
  position: SourcePosition;
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
  | "LIBREOFFICE_FONT_ENVIRONMENT_UNVERIFIED"
  | "RENDERER_PARITY_APPROXIMATE";

export type Diagnostic = {
  code: DiagnosticCode;
  severity: "warning";
  message: string;
  position?: SourcePosition;
  details?: Readonly<Record<string, JsonValue>>;
};

export type ErrorCode =
  | "INVALID_ARGUMENT"
  | "INTERNAL_ERROR"
  | "INVALID_CONFIG"
  | "INPUT_TOO_LARGE"
  | "DIFF_TOO_LARGE"
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
  | "NO_OFFICE_RENDERER"
  | "PROJECT_NOT_FOUND"
  | "PROJECT_INVALID"
  | "PROJECT_LOCKED"
  | "DOCUMENT_NOT_FOUND"
  | "DOCUMENT_EXISTS"
  | "PATH_OUTSIDE_PROJECT"
  | "REVISION_NOT_FOUND"
  | "REVISION_CONFLICT"
  | "WORKING_COPY_CONFLICT"
  | "PATCH_INVALID"
  | "PATCH_MISMATCH"
  | "PATCH_FAILED_VALIDATION"
  | "CHANGESET_INVALID"
  | "ANNOTATION_CONFLICT"
  | "REFERENCE_INVALID"
  | "RULE_PACK_INVALID"
  | "PAGINATION_DID_NOT_CONVERGE"
  | "DOCX_IMPORT_UNSUPPORTED"
  | "DOCX_REDLINE_UNSUPPORTED"
  | "DOCX_GENERATED_INVALID";

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
