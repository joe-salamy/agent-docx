import type { Diagnostic } from "../types.js";
import type { FilingKind, ProvenanceSource } from "../measurement.js";

export type BuiltInProfileId =
  | "us-district-conventional"
  | "frap-32"
  | "cand-civil";

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
