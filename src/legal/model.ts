import type { NormalizedSourceSegment } from "../markdown.js";
import type { SourcePosition } from "../types.js";
import type {
  BuiltInProfileId,
  FontSetInput,
  LayoutProfile,
} from "../layout/profile.js";
import type { FilingKind } from "../measurement.js";

export type RulePackId = "frap-32@2024-12-01" | "cand-civil@2026-05-01";
export type BlockId = `b_${string}`;
export type AnnotationId = `a_${string}`;
export type RevisionId = `sha256:${string}`;

export type Actor = {
  name: string;
  email?: string;
};

export type LitigationMetadata = {
  court: string;
  jurisdiction: string;
  caseName: string;
  docketNumber: string;
  documentTitle: string;
  filingDate?: string;
  parties: readonly { id: string; name: string; role: string }[];
  counsel: readonly {
    id: string;
    name: string;
    barNumber?: string;
    firm?: string;
    addressLines?: readonly string[];
    phone?: string;
    email?: string;
  }[];
  certificates: readonly (
    | {
        id: string;
        kind: "service";
        statement: string;
        servedOn: readonly string[];
        method: string;
        date?: string;
        signerCounselId: string;
      }
    | {
        id: string;
        kind: "compliance";
        basis: "words" | "monospaced-lines";
        signerCounselId: string;
      }
  )[];
};

export type DocumentChrome = {
  headers?: { default?: string; first?: string; even?: string };
  footers?: { default?: string; first?: string; even?: string };
  pageNumber?: {
    story: "header" | "footer";
    alignment: "left" | "center" | "right";
    format: "decimal" | "lower-roman" | "upper-roman";
    start: number;
  };
  lineNumbers?: {
    countBy: number;
    start: number;
    distanceTwips: number;
    restart: "continuous" | "new-page" | "new-section";
  };
};

export type AuthorityReference = {
  id: string;
  category: "cases" | "statutes" | "rules" | "other";
  short: string;
};

export type InlineRun = {
  text: string;
  bold: boolean;
  italic: boolean;
  strikethrough: boolean;
  literal: boolean;
  hardBreakAfter: boolean;
  link?: { target: string; title?: string };
  footnoteId?: string;
  referenceTarget?: BlockId;
  authority?: AuthorityReference;
};

export type BlockBase = {
  id: BlockId;
  position: SourcePosition;
  sourceText: string;
  segments: readonly NormalizedSourceSegment[];
};

export type InlineParagraph = {
  position: SourcePosition;
  sourceText: string;
  segments: readonly NormalizedSourceSegment[];
  runs: readonly InlineRun[];
};

export type LegalTableCell = {
  paragraphs: readonly InlineParagraph[];
  verticalAlign: "top" | "center" | "bottom";
};

export type FootnoteDefinition = BlockBase & {
  kind: "footnote";
  label: string;
  paragraphs: readonly InlineParagraph[];
};

export type LegalListItem = {
  position: SourcePosition;
  sourceText: string;
  segments: readonly NormalizedSourceSegment[];
  paragraphs: readonly InlineParagraph[];
  children: readonly LegalListBlock[];
};

export type LegalListBlock = BlockBase & {
  kind: "list";
  ordered: boolean;
  start: number | null;
  depth: number;
  items: readonly LegalListItem[];
};

export type LegalBlock =
  | (BlockBase & {
      kind: "paragraph";
      runs: readonly InlineRun[];
      footnoteRefs: readonly string[];
    })
  | (BlockBase & {
      kind: "blockquote";
      depth: number;
      runs: readonly InlineRun[];
      footnoteRefs: readonly string[];
    })
  | (BlockBase & {
      kind: "heading";
      level: 1 | 2 | 3 | 4 | 5 | 6;
      runs: readonly InlineRun[];
    })
  | (BlockBase & {
      kind: "numbered-paragraph";
      sequence: string;
      level: 1 | 2 | 3 | 4;
      runs: readonly InlineRun[];
    })
  | LegalListBlock
  | (BlockBase & {
      kind: "table";
      rows: readonly (readonly LegalTableCell[])[];
      align: readonly ("left" | "center" | "right" | null)[];
    })
  | (BlockBase & { kind: "caption" | "toc" | "toa" })
  | (BlockBase & { kind: "signature"; counselId: string })
  | (BlockBase & { kind: "certificate"; certificateId: string })
  | (BlockBase & {
      kind: "exhibit";
      exhibitId: string;
      label: string;
      source: string;
      blocks: readonly LegalBlock[];
    })
  | (BlockBase & {
      kind: "length-exclusion";
      exclusionKind:
        | "disclosure-statement"
        | "oral-argument-statement"
        | "statutory-addendum"
        | "proof-of-service"
        | "local-rule";
      citation?: string;
      blocks: readonly LegalBlock[];
    })
  | (BlockBase & {
      kind: "image";
      source: string;
      alt: string;
      widthTwips: number;
      heightTwips: number;
    })
  | (BlockBase & { kind: "pagebreak" | "thematic-break" })
  | (BlockBase & {
      kind: "sectionbreak";
      breakKind: "next-page" | "continuous";
      pageNumber?: {
        format: "decimal" | "lower-roman" | "upper-roman";
        start: number;
      };
    });

export type ReviewAnnotation = {
  id: AnnotationId;
  blockId: BlockId;
  range?: { start: number; end: number };
  author: Actor | null;
  createdAt: string | null;
  message: string;
  status: "open" | "resolved";
};

export type LegalDocument = {
  schemaVersion: 1;
  projectId: string;
  documentId: string;
  metadata: LitigationMetadata;
  chrome: DocumentChrome;
  blocks: readonly LegalBlock[];
  footnotes: readonly FootnoteDefinition[];
  annotations: readonly ReviewAnnotation[];
  assets: Readonly<
    Record<
      string,
      { sha256: `sha256:${string}`; mediaType: string; bytes: number }
    >
  >;
  source: { text: string; sha256: `sha256:${string}` };
};

export type LegalDocumentSpecification = {
  projectId?: string;
  documentId: string;
  profile: BuiltInProfileId | LayoutProfile;
  filingKind?: FilingKind;
  rulePack?: RulePackId;
  metadata: LitigationMetadata;
  chrome?: DocumentChrome;
  template?: Uint8Array;
  fontSet?: FontSetInput;
  assets?: Readonly<Record<string, { bytes: Uint8Array; mediaType: string }>>;
};

export type AddressableBlock = LegalBlock | FootnoteDefinition;
export type ContainerBlock =
  | LegalListBlock
  | Extract<LegalBlock, { kind: "exhibit" | "length-exclusion" }>;
export type LeafAddressableBlock = Exclude<AddressableBlock, ContainerBlock>;

export const emptyLitigationMetadata = (): LitigationMetadata => ({
  court: "",
  jurisdiction: "",
  caseName: "",
  docketNumber: "",
  documentTitle: "",
  parties: [],
  counsel: [],
  certificates: [],
});

export const isDocumentId = (value: string): boolean =>
  /^[a-z0-9][a-z0-9-]{0,63}$/.test(value);

export const isBlockId = (value: string): value is BlockId =>
  /^b_[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/.test(
    value,
  );

export const blockBookmark = (blockId: BlockId): string =>
  `adx_${blockId.slice(2).replaceAll("-", "")}`;
