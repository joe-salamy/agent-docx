import canonicalize from "canonicalize";
import {
  AlignmentType,
  Bookmark,
  BorderStyle,
  Document,
  ExternalHyperlink,
  Footer,
  FootnoteReferenceRun,
  Header,
  ImageRun,
  InternalHyperlink,
  LevelFormat,
  LineNumberRestartFormat,
  NumberFormat,
  Packer,
  PageNumber,
  Paragraph,
  SectionType,
  Table,
  TableCell,
  TableLayoutType,
  TableRow,
  Tab,
  TextRun,
  WidthType,
  type ISpacingProperties,
  type ParagraphChild,
} from "docx";
import type {
  InlineRun,
  NormalizedDocument,
  TableFlowBlock,
  TextFlowBlock,
} from "../markdown.js";
import { conservativeBodyBounds } from "../layout/body.js";
import { flowStyleFor } from "../layout/style.js";
import { tableColumnWidths } from "../layout/table.js";
import {
  blockBookmark,
  type DocumentChrome,
  type LegalDocument,
  type LitigationMetadata,
  type ReviewAnnotation,
} from "../legal/model.js";
import { lowerLegalDocument } from "../legal/lower.js";
import { AgentDocxError } from "../types.js";
import type { LayoutProfile, TextStyle } from "../layout/profile.js";
import type { SourcePosition } from "../types.js";
import type { ValidationResult } from "../legal/rules.js";
import type { ChangeSet, RevisionRecord } from "../revisions/types.js";
import type { LoadedFonts } from "../resolve.js";
import {
  decodeDocxXml,
  readDocxParts,
  repackDocxParts,
  sha256Hex,
} from "./package.js";
import { sanitizeXmlText } from "./xml-text.js";
import { sourcePartForRelationshipPart } from "./opc.js";

type TextContent = Pick<TextFlowBlock, "runs" | "image">;
type RichInlineRun = InlineRun & {
  strikethrough?: boolean;
  link?: { target: string; title?: string };
  referenceTarget?: string;
};
export type GenerateDocxOptions = {
  assets?: Readonly<Record<string, { bytes: Uint8Array; mediaType: string }>>;
  chrome?: DocumentChrome;
  metadata?: LitigationMetadata;
  /** Stabilized deterministic page count for PAGE/NUMPAGES chrome fields. */
  pageCount?: number;
  /** Loaded metric fonts; enables content-based table column allocation. */
  fonts?: LoadedFonts;
  revision?: RevisionRecord;
  changeSet?: ChangeSet;
  annotations?: readonly ReviewAnnotation[];
  validation?: ValidationResult;
  dependencies?: ReadonlyMap<
    string,
    {
      sha256: `sha256:${string}`;
      mediaType: string;
      bytes: Uint8Array;
    }
  >;
  semanticManifest?: Readonly<Record<string, unknown>>;
  createdAt?: string;
};
export type BodyParagraphManifestEntry = {
  id: string;
  index: number;
  position: SourcePosition;
  preview: string;
};
export type GeneratedDocx = {
  bytes: Uint8Array;
  bodyParagraphs: readonly BodyParagraphManifestEntry[];
};

const isLegalDocument = (
  value: NormalizedDocument | LegalDocument,
): value is LegalDocument =>
  "schemaVersion" in value &&
  "metadata" in value &&
  "chrome" in value &&
  "documentId" in value;

const dependencyAssets = (
  dependencies: GenerateDocxOptions["dependencies"],
): Readonly<Record<string, { bytes: Uint8Array; mediaType: string }>> => {
  if (!dependencies) return {};
  const assets: Record<string, { bytes: Uint8Array; mediaType: string }> = {};
  for (const [key, dependency] of dependencies) {
    if (!key.startsWith("asset/")) continue;
    const sha256 = sha256Hex(dependency.bytes);
    if (sha256 !== dependency.sha256)
      throw new AgentDocxError(
        "DOCX_GENERATED_INVALID",
        `Dependency hash does not match ${key}`,
      );
    assets[key.slice("asset/".length)] = {
      bytes: dependency.bytes,
      mediaType: dependency.mediaType,
    };
  }
  return assets;
};

const spacing = (style: TextStyle): ISpacingProperties =>
  style.lineSpacing.rule === "auto"
    ? {
        line: style.lineSpacing.numerator,
        lineRule: "auto",
        before: style.beforeTwips,
        after: style.afterTwips,
      }
    : {
        line: style.lineSpacing.twips,
        lineRule: style.lineSpacing.rule === "exact" ? "exact" : "atLeast",
        before: style.beforeTwips,
        after: style.afterTwips,
      };

export const paragraphOptions = (style: TextStyle, widowControl: boolean) => ({
  keepNext: style.keepWithNext,
  keepLines: style.keepLines,
  widowControl,
  spacing: spacing(style),
  indent: {
    left: style.leftIndentTwips,
    right: style.rightIndentTwips,
    firstLine: style.firstLineIndentTwips,
    hanging: style.hangingIndentTwips,
  },
});

const paragraphStyle = (
  id: string,
  name: string,
  style: TextStyle,
  fontFamily: string,
  options: {
    basedOn?: string;
    alignment?: (typeof AlignmentType)[keyof typeof AlignmentType];
    outlineLevel?: number;
  } = {},
) => ({
  id,
  name,
  ...(options.basedOn ? { basedOn: options.basedOn } : {}),
  paragraph: {
    spacing: spacing(style),
    indent: {
      left: style.leftIndentTwips,
      right: style.rightIndentTwips,
      firstLine: style.firstLineIndentTwips,
      hanging: style.hangingIndentTwips,
    },
    keepNext: style.keepWithNext,
    keepLines: style.keepLines,
    ...(options.alignment ? { alignment: options.alignment } : {}),
    ...(options.outlineLevel === undefined
      ? {}
      : { outlineLevel: options.outlineLevel }),
  },
  run: {
    font: fontFamily,
    size: style.fontSizeTwips / 10,
    bold: style.bold,
    italics: style.italic,
  },
});

export const nativeStyles = (profile: LayoutProfile) => {
  const font = profile.requestedFontFamily;
  const body = paragraphStyle(
    "AgentDocxBody",
    "AgentDocxBody",
    profile.body,
    font,
  );
  const headings = ([1, 2, 3, 4, 5, 6] as const).map((level) =>
    paragraphStyle(
      `AgentDocxHeading${level}`,
      `AgentDocxHeading${level}`,
      profile.headings[String(level) as "1"],
      font,
      { basedOn: "AgentDocxBody", outlineLevel: level - 1 },
    ),
  );
  const derivative = (
    id: string,
    style: TextStyle,
    options: Parameters<typeof paragraphStyle>[4] = {},
  ) =>
    paragraphStyle(id, id, style, font, {
      basedOn: "AgentDocxBody",
      ...options,
    });
  const single = (style: TextStyle): TextStyle => ({
    ...style,
    beforeTwips: 0,
    afterTwips: 0,
    lineSpacing: { rule: "auto", numerator: 240, denominator: 240 },
  });
  return [
    body,
    ...headings,
    derivative("AgentDocxBlockQuote", profile.blockquote),
    derivative("AgentDocxList", profile.list),
    ...([1, 2, 3, 4] as const).map((level) =>
      derivative(`AgentDocxNumbered${level}`, profile.list),
    ),
    derivative("AgentDocxFootnote", profile.footnote),
    derivative(
      "AgentDocxCaption",
      { ...profile.body, bold: true },
      { alignment: AlignmentType.CENTER },
    ),
    derivative("AgentDocxTOCHeading", profile.headings["1"]),
    derivative("AgentDocxTOCEntry", single(profile.body)),
    derivative("AgentDocxTOAHeading", profile.headings["1"]),
    derivative("AgentDocxTOAEntry", single(profile.body)),
    derivative("AgentDocxSignature", profile.body),
    derivative("AgentDocxCertificate", profile.body),
    derivative("AgentDocxHeader", single(profile.body)),
    derivative("AgentDocxFooter", single(profile.body)),
  ];
};

const styleFor = (
  block: TextFlowBlock,
  profile: LayoutProfile,
): { id: string; style: TextStyle } => {
  const resolved = flowStyleFor(block, profile);
  return { id: resolved.styleId, style: resolved.style };
};

const runChildren = (
  run: RichInlineRun,
  style: TextStyle,
  fontFamily: string,
  footnoteIds: ReadonlyMap<string, number>,
): ParagraphChild => {
  if (run.footnoteId !== undefined)
    return new FootnoteReferenceRun(footnoteIds.get(run.footnoteId)!);
  const trailingBreak = run.text.endsWith("\n");
  const sanitized = sanitizeXmlText(
    trailingBreak ? run.text.slice(0, -1) : run.text,
  );
  const children = sanitized.includes("\t")
    ? sanitized
        .split("\t")
        .flatMap((part, index) => [
          ...(index > 0 ? [new Tab()] : []),
          ...(part.length > 0 ? [part] : []),
        ])
    : undefined;
  const textRun = new TextRun({
    ...(children === undefined ? { text: sanitized } : { children }),
    font: run.literal ? "Courier New" : fontFamily,
    size: style.fontSizeTwips / 10,
    bold: style.bold || run.bold,
    italics: style.italic || run.italic,
    strike: run.strikethrough === true,
    ...(trailingBreak ? { break: 1 } : {}),
  });
  if (run.referenceTarget)
    return new InternalHyperlink({
      anchor: blockBookmark(
        run.referenceTarget as Parameters<typeof blockBookmark>[0],
      ),
      children: [textRun],
    });
  if (run.link)
    return new ExternalHyperlink({
      link: run.link.target,
      children: [textRun],
    });
  return textRun;
};

const textChildren = (
  block: TextContent,
  style: TextStyle,
  fontFamily: string,
  footnoteIds: ReadonlyMap<string, number>,
): ParagraphChild[] =>
  block.runs.map((run) =>
    runChildren(run as RichInlineRun, style, fontFamily, footnoteIds),
  );

const imageChildren = (
  block: TextContent,
  assets: GenerateDocxOptions["assets"],
): ParagraphChild[] | null => {
  if (!block.image) return null;
  const asset = assets?.[block.image.source];
  if (!asset)
    throw new AgentDocxError(
      "REFERENCE_INVALID",
      `Missing image asset: ${block.image.source}`,
    );
  const type =
    asset.mediaType === "image/png"
      ? "png"
      : asset.mediaType === "image/jpeg"
        ? "jpg"
        : null;
  if (!type)
    throw new AgentDocxError(
      "REFERENCE_INVALID",
      `Image must be PNG or JPEG: ${block.image.source}`,
    );
  return [
    new ImageRun({
      type,
      data: asset.bytes,
      transformation: {
        width: Math.max(1, Math.round(block.image.widthTwips / 15)),
        height: Math.max(1, Math.round(block.image.heightTwips / 15)),
      },
      altText: {
        name: sanitizeXmlText(block.image.source),
        description: sanitizeXmlText(block.image.alt),
        title: sanitizeXmlText(block.image.alt),
      },
    }),
  ];
};

const tableGridWidths = (
  block: TableFlowBlock,
  profile: LayoutProfile,
  usableWidth: number,
  fonts?: LoadedFonts,
): number[] => {
  if (fonts) return tableColumnWidths(block, profile, fonts, usableWidth);
  const columns = block.rows[0]?.length ?? 0;
  if (columns === 0)
    throw new AgentDocxError(
      "INVALID_LAYOUT",
      "A table must contain at least one column.",
    );
  const fixed =
    profile.table.cellPaddingTwips.left +
    profile.table.cellPaddingTwips.right +
    2 * profile.table.borderTwips;
  const styleFloor = (style: TextStyle) =>
    Math.max(0, style.leftIndentTwips) +
    Math.max(0, style.rightIndentTwips) +
    Math.max(0, style.firstLineIndentTwips - style.hangingIndentTwips);
  const floor =
    fixed +
    Math.max(styleFloor(profile.table.header), styleFloor(profile.table.body)) +
    1;
  const widths = Array<number>(columns).fill(floor);
  let remaining =
    usableWidth - widths.reduce((total, value) => total + value, 0);
  for (let index = 0; remaining > 0; index = (index + 1) % columns) {
    widths[index]!++;
    remaining--;
  }
  return widths;
};

export const alignmentFor = (
  alignment: "left" | "center" | "right" | null,
): (typeof AlignmentType)[keyof typeof AlignmentType] =>
  alignment === "center"
    ? AlignmentType.CENTER
    : alignment === "right"
      ? AlignmentType.RIGHT
      : AlignmentType.LEFT;

const templateTokens = (
  value: string,
  metadata: LitigationMetadata | undefined,
  style: TextStyle,
  fontFamily: string,
): ParagraphChild[] => {
  const safeValue = sanitizeXmlText(value);
  const values: Record<string, string> = {
    caseName: sanitizeXmlText(metadata?.caseName ?? ""),
    docketNumber: sanitizeXmlText(metadata?.docketNumber ?? ""),
    documentTitle: sanitizeXmlText(metadata?.documentTitle ?? ""),
  };
  const children: ParagraphChild[] = [];
  let index = 0;
  for (const match of safeValue.matchAll(
    /\{\{(caseName|docketNumber|documentTitle|page|pages)\}\}/g,
  )) {
    if (match.index! > index)
      children.push(
        new TextRun({
          text: safeValue.slice(index, match.index),
          font: fontFamily,
          size: style.fontSizeTwips / 10,
        }),
      );
    const token = match[1]!;
    if (token === "page")
      children.push(new TextRun({ children: [PageNumber.CURRENT] }));
    else if (token === "pages")
      children.push(new TextRun({ children: [PageNumber.TOTAL_PAGES] }));
    else
      children.push(
        new TextRun({
          text: values[token]!,
          font: fontFamily,
          size: style.fontSizeTwips / 10,
        }),
      );
    index = match.index! + match[0].length;
  }
  if (index < safeValue.length || children.length === 0)
    children.push(
      new TextRun({
        text: safeValue.slice(index),
        font: fontFamily,
        size: style.fontSizeTwips / 10,
      }),
    );
  return children;
};

export const chromePart = (
  value: string | undefined,
  styleId: "AgentDocxHeader" | "AgentDocxFooter",
  profile: LayoutProfile,
  metadata: LitigationMetadata | undefined,
  alignment: (typeof AlignmentType)[keyof typeof AlignmentType] = AlignmentType.LEFT,
): Paragraph | null =>
  value === undefined
    ? null
    : new Paragraph({
        style: styleId,
        alignment,
        children: templateTokens(
          value,
          metadata,
          profile.body,
          profile.requestedFontFamily,
        ),
      });

const numberFormat = (
  format: "decimal" | "lower-roman" | "upper-roman",
): (typeof NumberFormat)[keyof typeof NumberFormat] =>
  format === "lower-roman"
    ? NumberFormat.LOWER_ROMAN
    : format === "upper-roman"
      ? NumberFormat.UPPER_ROMAN
      : NumberFormat.DECIMAL;

const lineRestart = (
  restart: "continuous" | "new-page" | "new-section",
): (typeof LineNumberRestartFormat)[keyof typeof LineNumberRestartFormat] =>
  restart === "new-page"
    ? LineNumberRestartFormat.NEW_PAGE
    : restart === "new-section"
      ? LineNumberRestartFormat.NEW_SECTION
      : LineNumberRestartFormat.CONTINUOUS;

export const numbering = (
  profile: LayoutProfile,
  orderedListStarts: readonly number[] = [],
) => {
  const orderedLevels = (start?: number) =>
    [0, 1, 2, 3].map((level) => ({
      level,
      format: LevelFormat.DECIMAL,
      text: `%${level + 1}.`,
      alignment: AlignmentType.LEFT,
      ...(start !== undefined && level === 0 ? { start } : {}),
      style: {
        paragraph: {
          indent: {
            left: profile.list.leftIndentTwips + (level + 1) * 360,
            hanging: Math.max(180, profile.list.hangingIndentTwips || 360),
          },
        },
      },
    }));
  return {
    config: [
      {
        reference: "AgentDocxOrderedList",
        levels: orderedLevels(),
      },
      ...orderedListStarts.map((start) => ({
        reference: `AgentDocxOrderedListFrom${start}`,
        levels: orderedLevels(start),
      })),
      {
        reference: "AgentDocxBulletList",
        levels: [0, 1, 2, 3].map((level) => ({
          level,
          format: LevelFormat.BULLET,
          text: "•",
          alignment: AlignmentType.LEFT,
          style: {
            paragraph: {
              indent: {
                left: profile.list.leftIndentTwips + (level + 1) * 360,
                hanging: Math.max(180, profile.list.hangingIndentTwips || 360),
              },
            },
          },
        })),
      },
      {
        reference: "AgentDocxNumberedParagraph",
        levels: [0, 1, 2, 3].map((level) => ({
          level,
          format: LevelFormat.DECIMAL,
          text: `%${level + 1}.`,
          alignment: AlignmentType.LEFT,
          style: {
            paragraph: {
              indent: {
                left: profile.list.leftIndentTwips + (level + 1) * 360,
                hanging: Math.max(180, profile.list.hangingIndentTwips || 360),
              },
            },
          },
        })),
      },
    ],
  };
};

const toTable = (
  block: TableFlowBlock,
  profile: LayoutProfile,
  usableWidth: number,
  footnoteIds: ReadonlyMap<string, number>,
  fonts?: LoadedFonts,
): Table => {
  const gridWidths = tableGridWidths(block, profile, usableWidth, fonts);
  const borderSize =
    profile.table.borderTwips === 0
      ? 0
      : Math.max(1, Math.floor((profile.table.borderTwips * 2) / 5 + 0.5));
  const tableBorder = {
    style:
      profile.table.borderTwips === 0 ? BorderStyle.NONE : BorderStyle.SINGLE,
    size: borderSize,
    color: "000000",
  };
  return new Table({
    width: { size: usableWidth, type: WidthType.DXA },
    columnWidths: gridWidths,
    layout: TableLayoutType.FIXED,
    borders: {
      top: tableBorder,
      bottom: tableBorder,
      left: tableBorder,
      right: tableBorder,
      insideHorizontal: tableBorder,
      insideVertical: tableBorder,
    },
    margins: {
      marginUnitType: WidthType.DXA,
      top: profile.table.cellPaddingTwips.top,
      right: profile.table.cellPaddingTwips.right,
      bottom: profile.table.cellPaddingTwips.bottom,
      left: profile.table.cellPaddingTwips.left,
    },
    rows: block.rows.map((row, rowIndex) => {
      const style = rowIndex === 0 ? profile.table.header : profile.table.body;
      return new TableRow({
        cantSplit: false,
        tableHeader: rowIndex === 0 && profile.table.repeatHeader,
        children: row.slice(0, gridWidths.length).map(
          (cell, columnIndex) =>
            new TableCell({
              width: { size: gridWidths[columnIndex]!, type: WidthType.DXA },
              margins: {
                marginUnitType: WidthType.DXA,
                top: profile.table.cellPaddingTwips.top,
                right: profile.table.cellPaddingTwips.right,
                bottom: profile.table.cellPaddingTwips.bottom,
                left: profile.table.cellPaddingTwips.left,
              },
              borders: {
                top: tableBorder,
                bottom: tableBorder,
                left: tableBorder,
                right: tableBorder,
              },
              children: [
                new Paragraph({
                  style: "AgentDocxBody",
                  ...paragraphOptions(
                    style,
                    profile.pagination.widowOrphanControl,
                  ),
                  alignment: alignmentFor(cell.alignment),
                  children: textChildren(
                    cell,
                    style,
                    profile.requestedFontFamily,
                    footnoteIds,
                  ),
                }),
              ],
            }),
        ),
      });
    }),
  });
};

const toBorderlessLegalTable = (
  styleId: string,
  style: TextStyle,
  profile: LayoutProfile,
  usableWidth: number,
  alignment: (typeof AlignmentType)[keyof typeof AlignmentType],
  children: readonly ParagraphChild[],
): Table => {
  const border = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
  return new Table({
    width: { size: usableWidth, type: WidthType.DXA },
    columnWidths: [usableWidth],
    layout: TableLayoutType.FIXED,
    borders: {
      top: border,
      bottom: border,
      left: border,
      right: border,
      insideHorizontal: border,
      insideVertical: border,
    },
    rows: [
      new TableRow({
        cantSplit: true,
        children: [
          new TableCell({
            width: { size: usableWidth, type: WidthType.DXA },
            borders: {
              top: border,
              bottom: border,
              left: border,
              right: border,
            },
            children: [
              new Paragraph({
                style: styleId,
                ...paragraphOptions(
                  style,
                  profile.pagination.widowOrphanControl,
                ),
                alignment,
                children,
              }),
            ],
          }),
        ],
      }),
    ],
  });
};

const escapeXmlText = (value: string): string =>
  sanitizeXmlText(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");

const appendBefore = (
  xml: string,
  closingTag: string,
  fragment: string,
): string => {
  const index = xml.lastIndexOf(closingTag);
  if (index === -1)
    throw new AgentDocxError(
      "DOCX_GENERATED_INVALID",
      `Generated package is missing ${closingTag}`,
    );
  return `${xml.slice(0, index)}${fragment}${xml.slice(index)}`;
};

const normalizedTimestamp = (requested: string | undefined): string => {
  if (requested === undefined) return "1980-01-01T00:00:00.000Z";
  const parsed = new Date(requested);
  if (Number.isNaN(parsed.getTime()))
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      "Generated DOCX timestamp must be an ISO-8601 date",
    );
  return parsed.toISOString();
};

const normalizeCoreProperties = (xml: string, timestamp: string): string =>
  xml
    .replace(
      /(<dcterms:created\b[^>]*>).*?(<\/dcterms:created>)/s,
      `$1${timestamp}$2`,
    )
    .replace(
      /(<dcterms:modified\b[^>]*>).*?(<\/dcterms:modified>)/s,
      `$1${timestamp}$2`,
    );

const normalizeRelationshipIds = (parts: Map<string, Uint8Array>): void => {
  const relationshipParts = [...parts.keys()]
    .filter((path) => path.endsWith(".rels"))
    .sort((left, right) => (left < right ? -1 : left > right ? 1 : 0));
  for (const relationshipPart of relationshipParts) {
    const sourcePart = sourcePartForRelationshipPart(relationshipPart);
    const original = decodeDocxXml(parts.get(relationshipPart)!);
    const ids = new Map<string, string>();
    let nextId = 1;
    const rewritten = original.replace(
      /(<Relationship\b[^>]*\bId=")([^"]+)(")/g,
      (_match, prefix: string, id: string, suffix: string) => {
        const replacement = `rId${nextId++}`;
        ids.set(id, replacement);
        return `${prefix}${replacement}${suffix}`;
      },
    );
    parts.set(relationshipPart, new TextEncoder().encode(rewritten));
    if (sourcePart.length === 0 || ids.size === 0) continue;
    const source = parts.get(sourcePart);
    if (!source) continue;
    const sourceXml = decodeDocxXml(source).replace(
      /((?:r:(?:id|embed|link))=")([^"]+)(")/g,
      (_match, prefix: string, id: string, suffix: string) =>
        ids.has(id) ? `${prefix}${ids.get(id)!}${suffix}` : _match,
    );
    parts.set(sourcePart, new TextEncoder().encode(sourceXml));
  }
};

export const normalizeGeneratedPackage = async (
  bytes: Uint8Array,
  createdAt: string | undefined,
): Promise<Uint8Array> => {
  const parts = new Map(await readDocxParts(bytes));
  normalizeRelationshipIds(parts);
  const core = parts.get("docProps/core.xml");
  if (core)
    parts.set(
      "docProps/core.xml",
      new TextEncoder().encode(
        normalizeCoreProperties(
          decodeDocxXml(core),
          normalizedTimestamp(createdAt),
        ),
      ),
    );
  return repackDocxParts(parts);
};

export const addSemanticManifest = async (
  bytes: Uint8Array,
  manifest: Readonly<Record<string, unknown>>,
  createdAt: string | undefined,
): Promise<Uint8Array> => {
  const parts = new Map(await readDocxParts(bytes));
  normalizeRelationshipIds(parts);
  const relationshipPart = "word/_rels/document.xml.rels";
  const rels = parts.get(relationshipPart);
  const contentTypes = parts.get("[Content_Types].xml");
  if (!rels || !contentTypes)
    throw new AgentDocxError(
      "DOCX_GENERATED_INVALID",
      "Generated package lacks its main relationship or content-types part",
    );
  const manifestPath = "customXml/itemAgentDocx.xml";
  if (parts.has(manifestPath))
    throw new AgentDocxError(
      "DOCX_GENERATED_INVALID",
      "Generated package already contains the agent semantic manifest path",
    );
  const relsXml = decodeDocxXml(rels);
  let relationshipId = "rIdAgentDocxSemantic";
  let suffix = 1;
  while (new RegExp(`\\bId="${relationshipId}"`).test(relsXml))
    relationshipId = `rIdAgentDocxSemantic${suffix++}`;
  parts.set(
    relationshipPart,
    new TextEncoder().encode(
      appendBefore(
        relsXml,
        "</Relationships>",
        `<Relationship Id="${relationshipId}" Type="https://agent-docx.dev/relationships/semantic-manifest" Target="../${manifestPath}"/>`,
      ),
    ),
  );
  const contentXml = decodeDocxXml(contentTypes);
  parts.set(
    "[Content_Types].xml",
    new TextEncoder().encode(
      appendBefore(
        contentXml,
        "</Types>",
        `<Override PartName="/${manifestPath}" ContentType="application/xml"/>`,
      ),
    ),
  );
  const payload = canonicalize(manifest);
  if (payload === undefined)
    throw new AgentDocxError(
      "DOCX_GENERATED_INVALID",
      "Semantic manifest cannot be canonicalized",
    );
  parts.set(
    manifestPath,
    new TextEncoder().encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><agent-docx xmlns="https://agent-docx.dev/semantic-manifest/v1"><payload>${escapeXmlText(payload)}</payload></agent-docx>`,
    ),
  );
  const corePath = "docProps/core.xml";
  const core = parts.get(corePath);
  if (core)
    parts.set(
      corePath,
      new TextEncoder().encode(
        normalizeCoreProperties(
          decodeDocxXml(core),
          normalizedTimestamp(createdAt),
        ),
      ),
    );
  return repackDocxParts(parts);
};

export const createNativeDocumentChrome = (
  profile: LayoutProfile,
  chrome: DocumentChrome = {},
  metadata: LitigationMetadata | undefined,
  pageCount: number | undefined,
) => {
  const bodyBounds = conservativeBodyBounds(
    profile,
    chrome,
    pageCount === undefined ? 1 : String(pageCount).length,
  );
  const headerTemplates = chrome.headers ?? {};
  const footerTemplates = chrome.footers ?? {};
  const pageNumberTemplate = chrome.pageNumber;
  const pageNumberParagraph = pageNumberTemplate
    ? new Paragraph({
        style:
          pageNumberTemplate.story === "header"
            ? "AgentDocxHeader"
            : "AgentDocxFooter",
        alignment: alignmentFor(pageNumberTemplate.alignment),
        children: [new TextRun({ children: [PageNumber.CURRENT] })],
      })
    : null;
  const makeHeader = (value: string | undefined) => {
    const paragraph = chromePart(value, "AgentDocxHeader", profile, metadata);
    const number =
      pageNumberTemplate?.story === "header" ? pageNumberParagraph : null;
    return paragraph || number
      ? new Header({
          children: [paragraph, number].filter(
            (item): item is Paragraph => item !== null,
          ),
        })
      : undefined;
  };
  const makeFooter = (value: string | undefined) => {
    const paragraph = chromePart(value, "AgentDocxFooter", profile, metadata);
    const number =
      pageNumberTemplate?.story === "footer" ? pageNumberParagraph : null;
    return paragraph || number
      ? new Footer({
          children: [paragraph, number].filter(
            (item): item is Paragraph => item !== null,
          ),
        })
      : undefined;
  };
  const defaultHeader = makeHeader(headerTemplates.default);
  const firstHeader = makeHeader(headerTemplates.first);
  const evenHeader = makeHeader(headerTemplates.even);
  const headers = {
    ...(defaultHeader ? { default: defaultHeader } : {}),
    ...(firstHeader ? { first: firstHeader } : {}),
    ...(evenHeader ? { even: evenHeader } : {}),
  };
  const defaultFooter = makeFooter(footerTemplates.default);
  const firstFooter = makeFooter(footerTemplates.first);
  const evenFooter = makeFooter(footerTemplates.even);
  const footers = {
    ...(defaultFooter ? { default: defaultFooter } : {}),
    ...(firstFooter ? { first: firstFooter } : {}),
    ...(evenFooter ? { even: evenFooter } : {}),
  };
  return {
    bodyBounds,
    headers,
    footers,
    evenAndOddHeaderAndFooters:
      headerTemplates.even !== undefined || footerTemplates.even !== undefined,
    titlePage:
      headerTemplates.first !== undefined ||
      footerTemplates.first !== undefined,
  };
};

export const nativeSectionProperties = (
  profile: LayoutProfile,
  chrome: DocumentChrome,
  nativeChrome: ReturnType<typeof createNativeDocumentChrome>,
  pageNumber?: Pick<
    NonNullable<DocumentChrome["pageNumber"]>,
    "format" | "start"
  >,
  includeGlobalPageNumber = true,
) => ({
  page: {
    size: {
      width: profile.page.widthTwips,
      height: profile.page.heightTwips,
    },
    margin: {
      top: nativeChrome.bodyBounds.bodyTopTwips,
      right: profile.page.marginsTwips.right,
      bottom:
        profile.page.heightTwips - nativeChrome.bodyBounds.bodyBottomTwips,
      left: profile.page.marginsTwips.left,
      header: profile.page.headerTwips,
      footer: profile.page.footerTwips,
      gutter: profile.page.gutterTwips,
    },
    ...(pageNumber || (includeGlobalPageNumber ? chrome.pageNumber : undefined)
      ? {
          pageNumbers: {
            start: (pageNumber ?? chrome.pageNumber)!.start,
            formatType: numberFormat((pageNumber ?? chrome.pageNumber)!.format),
          },
        }
      : {}),
  },
  ...(chrome.lineNumbers
    ? {
        lineNumbers: {
          countBy: chrome.lineNumbers.countBy,
          start: chrome.lineNumbers.start,
          distance: chrome.lineNumbers.distanceTwips,
          restart: lineRestart(chrome.lineNumbers.restart),
        },
      }
    : {}),
  titlePage: nativeChrome.titlePage,
});

export async function generateDocx(
  input: NormalizedDocument | LegalDocument,
  profile: LayoutProfile,
  options: GenerateDocxOptions = {},
): Promise<GeneratedDocx> {
  const legalDocument = isLegalDocument(input) ? input : null;
  if (legalDocument)
    options = {
      ...options,
      chrome: options.chrome ?? legalDocument.chrome,
      metadata: options.metadata ?? legalDocument.metadata,
    };
  const assets = dependencyAssets(options.dependencies);
  options = {
    ...options,
    assets: { ...assets, ...options.assets },
    ...((options.createdAt ?? options.revision?.createdAt)
      ? { createdAt: options.createdAt ?? options.revision!.createdAt }
      : {}),
  };
  const flow: NormalizedDocument = legalDocument
    ? lowerLegalDocument(legalDocument)
    : (input as NormalizedDocument);
  const footnoteIds = new Map<string, number>();
  let nextId = 1;
  for (const id of flow.footnotes.keys()) footnoteIds.set(id, nextId++);
  const footnotes: Record<string, { children: readonly Paragraph[] }> = {};
  for (const [id, definition] of flow.footnotes) {
    const numeric = footnoteIds.get(id)!;
    footnotes[String(numeric)] = {
      children: definition.blocks.map((block, blockIndex) => {
        const style: TextStyle = {
          ...profile.footnote,
          beforeTwips: blockIndex === 0 ? 0 : profile.footnote.beforeTwips,
          afterTwips:
            blockIndex === definition.blocks.length - 1
              ? 0
              : profile.footnote.afterTwips,
          keepWithNext:
            blockIndex < definition.blocks.length - 1 &&
            profile.footnote.keepWithNext,
        };
        return new Paragraph({
          style: "AgentDocxFootnote",
          ...paragraphOptions(style, profile.pagination.widowOrphanControl),
          children: textChildren(
            block,
            style,
            profile.requestedFontFamily,
            footnoteIds,
          ),
        });
      }),
    };
  }
  const usableWidth =
    profile.page.widthTwips -
    profile.page.marginsTwips.left -
    profile.page.marginsTwips.right -
    profile.page.gutterTwips;
  const bodyParagraphs: BodyParagraphManifestEntry[] = [];
  let paragraphIndex = 0;
  type SectionContent = {
    children: Array<Paragraph | Table>;
    type?: (typeof SectionType)[keyof typeof SectionType];
    pageNumber?: {
      format: "decimal" | "lower-roman" | "upper-roman";
      start: number;
    };
  };
  const sections: SectionContent[] = [];
  let current: SectionContent = { children: [] };
  const finishSection = (): void => {
    if (current.children.length === 0)
      current.children.push(new Paragraph({ children: [] }));
    sections.push(current);
    current = { children: [] };
  };
  let pendingPageBreak = false;

  for (const rawBlock of flow.blocks) {
    const sectionBreak =
      rawBlock.kind === "pagebreak" ? rawBlock.sectionBreak : undefined;
    if (sectionBreak) {
      finishSection();
      current.type =
        sectionBreak.kind === "continuous"
          ? SectionType.CONTINUOUS
          : SectionType.NEXT_PAGE;
      if (sectionBreak.pageNumber !== undefined)
        current.pageNumber = sectionBreak.pageNumber;
      continue;
    }
    if (rawBlock.kind === "pagebreak") {
      pendingPageBreak = true;
      continue;
    }
    if (rawBlock.kind === "thematic-break") {
      current.children.push(
        new Paragraph({
          keepNext: profile.thematicBreak.keepWithNext,
          keepLines: true,
          ...(pendingPageBreak ? { pageBreakBefore: true } : {}),
          spacing: {
            before: profile.thematicBreak.beforeTwips,
            after: profile.thematicBreak.afterTwips,
            line: profile.thematicBreak.thicknessTwips,
            lineRule: "exact",
          },
          indent: { left: 0, right: 0, firstLine: 0, hanging: 0 },
          border: {
            bottom: {
              style: BorderStyle.SINGLE,
              color: "000000",
              size: Math.max(
                1,
                Math.floor(
                  (profile.thematicBreak.thicknessTwips * 2) / 5 + 0.5,
                ),
              ),
            },
          },
          children: [],
        }),
      );
      pendingPageBreak = false;
      continue;
    }
    if (rawBlock.kind === "table") {
      if (pendingPageBreak)
        current.children.push(
          new Paragraph({ pageBreakBefore: true, children: [] }),
        );
      current.children.push(
        toTable(rawBlock, profile, usableWidth, footnoteIds, options.fonts),
      );
      pendingPageBreak = false;
      continue;
    }
    const block = rawBlock;
    const resolved = styleFor(block, profile);
    const legalBlockId = block.legalBlockId;
    const id = legalBlockId
      ? blockBookmark(legalBlockId as Parameters<typeof blockBookmark>[0])
      : `adx_body_${String(paragraphIndex).padStart(6, "0")}`;
    const manifest = {
      id,
      index: paragraphIndex++,
      position: block.position,
      preview: block.normalizedText.replace(/\s+/g, " ").trim().slice(0, 80),
    };
    bodyParagraphs.push(manifest);
    const blockChildren =
      imageChildren(block, options.assets) ??
      textChildren(
        block,
        resolved.style,
        profile.requestedFontFamily,
        footnoteIds,
      );
    const bookmark = new Bookmark({ id, children: blockChildren });
    const numberingOptions =
      block.kind === "list"
        ? block.legalKind === "numbered-paragraph"
          ? {
              reference: "AgentDocxNumberedParagraph",
              level: Math.min(3, Math.max(0, block.numberedLevel ?? 0)),
            }
          : {
              reference: block.listOrdered
                ? block.listStart !== undefined
                  ? `AgentDocxOrderedListFrom${block.listStart}`
                  : "AgentDocxOrderedList"
                : "AgentDocxBulletList",
              level: Math.min(3, Math.max(0, block.listLevel ?? 0)),
            }
        : undefined;
    const legalTable =
      block.legalKind === "caption" ||
      block.legalKind === "signature" ||
      block.legalKind === "certificate";
    if (legalTable) {
      if (pendingPageBreak)
        current.children.push(
          new Paragraph({ pageBreakBefore: true, children: [] }),
        );
      current.children.push(
        toBorderlessLegalTable(
          resolved.id,
          resolved.style,
          profile,
          usableWidth,
          block.legalKind === "caption"
            ? AlignmentType.CENTER
            : AlignmentType.LEFT,
          [bookmark],
        ),
      );
    } else
      current.children.push(
        new Paragraph({
          style: resolved.id,
          ...paragraphOptions(
            resolved.style,
            profile.pagination.widowOrphanControl,
          ),
          ...(pendingPageBreak ? { pageBreakBefore: true } : {}),
          ...(block.legalKind === "caption"
            ? { alignment: AlignmentType.CENTER }
            : {}),
          ...(numberingOptions ? { numbering: numberingOptions } : {}),
          children: [bookmark],
        }),
      );
    pendingPageBreak = false;
  }
  finishSection();

  const chrome = options.chrome ?? {};
  const nativeChrome = createNativeDocumentChrome(
    profile,
    chrome,
    options.metadata,
    options.pageCount,
  );
  const orderedListStarts = [
    ...new Set(
      flow.blocks.flatMap((block) =>
        block.kind === "pagebreak" ||
        block.kind === "table" ||
        block.kind === "thematic-break"
          ? []
          : block.listStart !== undefined
            ? [block.listStart]
            : [],
      ),
    ),
  ].sort((left, right) => left - right);
  const document = new Document({
    footnotes,
    styles: { paragraphStyles: nativeStyles(profile) },
    numbering: numbering(profile, orderedListStarts),
    features: { updateFields: true },
    evenAndOddHeaderAndFooters: nativeChrome.evenAndOddHeaderAndFooters,
    sections: sections.map((section, sectionIndex) => ({
      properties: {
        ...(section.type ? { type: section.type } : {}),
        ...nativeSectionProperties(
          profile,
          chrome,
          nativeChrome,
          section.pageNumber,
          sectionIndex === 0,
        ),
      },
      ...(Object.keys(nativeChrome.headers).length > 0
        ? { headers: nativeChrome.headers }
        : {}),
      ...(Object.keys(nativeChrome.footers).length > 0
        ? { footers: nativeChrome.footers }
        : {}),
      children: section.children,
    })),
  });
  const packed = await Packer.toBuffer(document);
  const emittedBlocks =
    options.semanticManifest &&
    Array.isArray(options.semanticManifest.emittedBlocks)
      ? options.semanticManifest.emittedBlocks
      : bodyParagraphs.map(({ id, index }) => ({
          bookmark: id,
          index,
        }));
  const bytes = options.semanticManifest
    ? await addSemanticManifest(
        packed,
        {
          ...options.semanticManifest,
          emittedBlocks,
        },
        options.createdAt,
      )
    : await normalizeGeneratedPackage(packed, options.createdAt);
  return { bytes, bodyParagraphs };
}
