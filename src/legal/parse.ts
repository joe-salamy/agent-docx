import { createHash } from "node:crypto";
import remarkDirective from "remark-directive";
import remarkGfm from "remark-gfm";
import remarkParse from "remark-parse";
import { unified } from "unified";
import type { NormalizedSourceSegment } from "../markdown.js";
import { AgentDocxError, type SourcePosition } from "../types.js";
import { isSafeRelativePath } from "../path-util.js";
import {
  type BlockId,
  type DocumentChrome,
  type FootnoteDefinition,
  type InlineParagraph,
  type InlineRun,
  type LegalBlock,
  type LegalDocument,
  type LegalListBlock,
  type LegalListItem,
  type LegalTableCell,
  type LitigationMetadata,
  emptyLitigationMetadata,
  isBlockId,
  isDocumentId,
} from "./model.js";

type MarkdownPoint = { line: number; column: number; offset?: number };
type MarkdownNode = {
  type: string;
  value?: string;
  depth?: number;
  ordered?: boolean;
  start?: number;
  identifier?: string;
  label?: string;
  url?: string;
  title?: string | null;
  name?: string;
  attributes?: Record<string, string | null | undefined>;
  align?: Array<"left" | "center" | "right" | null>;
  children?: MarkdownNode[];
  position?: { start: MarkdownPoint; end: MarkdownPoint };
};

export type LegalAssetInput = {
  bytes: Uint8Array;
  mediaType: string;
};

export type ParseLegalMarkdownOptions = {
  projectId?: string;
  documentId: string;
  metadata?: LitigationMetadata;
  chrome?: DocumentChrome;
  assets?: Readonly<Record<string, LegalAssetInput>>;
  annotations?: LegalDocument["annotations"];
  requireMarkers?: boolean;
  exactAssets?: boolean;
};

export type ParsedLegalMarkdown = {
  document: LegalDocument;
  missingMarkers: readonly { offset: number; id: BlockId }[];
};

type Marker = { offset: number; end: number; id: BlockId };
type MarkerIndex = {
  byOffset: ReadonlyMap<number, Marker>;
  sorted: readonly Marker[];
};
type InlineResult = {
  runs: InlineRun[];
  text: string;
  segments: NormalizedSourceSegment[];
};
type ParseContext = {
  markdown: string;
  options: ParseLegalMarkdownOptions;
  markers: MarkerIndex;
  usedMarkers: Set<number>;
  missingMarkers: Array<{ offset: number; id: BlockId }>;
  assets: Map<string, LegalAssetInput>;
};

const sha256 = (value: string | Uint8Array): `sha256:${string}` =>
  `sha256:${createHash("sha256").update(value).digest("hex")}`;

const positionOf = (node: MarkdownNode): SourcePosition => {
  if (!node.position)
    throw new AgentDocxError(
      "UNSUPPORTED_MARKDOWN",
      "Markdown node has no source position",
    );
  return {
    start: {
      line: node.position.start.line,
      column: node.position.start.column,
      offset: node.position.start.offset ?? 0,
    },
    end: {
      line: node.position.end.line,
      column: node.position.end.column,
      offset: node.position.end.offset ?? 0,
    },
  };
};
const isXml10CodePoint = (value: number): boolean =>
  value === 0x09 ||
  value === 0x0a ||
  value === 0x0d ||
  (value >= 0x20 && value <= 0xd7ff) ||
  (value >= 0xe000 && value <= 0xfffd) ||
  (value >= 0x10000 && value <= 0x10ffff);
const xml10CodePointLabel = (value: number): string =>
  `U+${value.toString(16).toUpperCase().padStart(4, "0")}`;
const sourcePointAtOffset = (
  markdown: string,
  offset: number,
): SourcePosition["start"] => {
  let line = 1;
  let column = 1;
  for (let index = 0; index < offset; index++) {
    if (markdown.charCodeAt(index) === 0x0a) {
      line++;
      column = 1;
    } else column++;
  }
  return { line, column, offset };
};
const assertXml10Legal = (markdown: string): void => {
  for (let offset = 0; offset < markdown.length; ) {
    const first = markdown.charCodeAt(offset);
    let codePoint = first;
    let width = 1;
    if (first >= 0xd800 && first <= 0xdbff) {
      const second = markdown.charCodeAt(offset + 1);
      if (second >= 0xdc00 && second <= 0xdfff) {
        codePoint = 0x10000 + ((first - 0xd800) << 10) + (second - 0xdc00);
        width = 2;
      }
    }
    if (
      first >= 0xd800 &&
      first <= 0xdfff &&
      !(width === 2 && codePoint > 0xffff)
    ) {
      const label = xml10CodePointLabel(codePoint);
      throw new AgentDocxError(
        "UNSUPPORTED_MARKDOWN",
        `Markdown contains XML-1.0-illegal code point ${label}`,
        {
          position: {
            start: sourcePointAtOffset(markdown, offset),
            end: sourcePointAtOffset(markdown, offset + width),
          } as unknown as Record<string, never>,
          codePoint: label,
        },
      );
    }
    if (!isXml10CodePoint(codePoint)) {
      const label = xml10CodePointLabel(codePoint);
      throw new AgentDocxError(
        "UNSUPPORTED_MARKDOWN",
        `Markdown contains XML-1.0-illegal code point ${label}`,
        {
          position: {
            start: sourcePointAtOffset(markdown, offset),
            end: sourcePointAtOffset(markdown, offset + width),
          } as unknown as Record<string, never>,
          codePoint: label,
        },
      );
    }
    offset += width;
  }
};

const sourceTextOf = (node: MarkdownNode, markdown: string): string => {
  const position = positionOf(node);
  return markdown.slice(position.start.offset, position.end.offset);
};

const errorAt = (
  code: "UNSUPPORTED_MARKDOWN" | "REFERENCE_INVALID",
  message: string,
  node: MarkdownNode,
): never => {
  throw new AgentDocxError(code, message, {
    position: positionOf(node) as unknown as Record<string, never>,
  });
};

const deterministicBlockId = (
  documentId: string,
  kind: string,
  offset: number,
  sourceText: string,
): BlockId => {
  const hex = createHash("sha256")
    .update(documentId)
    .update("\0")
    .update(kind)
    .update("\0")
    .update(String(offset))
    .update("\0")
    .update(sourceText)
    .digest("hex")
    .slice(0, 32)
    .split("");
  hex[12] = "4";
  hex[16] = "8";
  const value = hex.join("");
  return `b_${value.slice(0, 8)}-${value.slice(8, 12)}-${value.slice(
    12,
    16,
  )}-${value.slice(16, 20)}-${value.slice(20)}`;
};

const markerPattern =
  /^[ \t]*<!--[ \t]*agent-docx:block[ \t]+id="(b_[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12})"[ \t]*-->(?:\r?\n|$)/gm;

const markerHtmlPattern =
  /^[ \t]*<!--[ \t]*agent-docx:block[ \t]+id="(b_[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12})"[ \t]*-->$/;

const MAX_NESTING_DEPTH = 100;
const normalizeFootnoteLabel = (value: string): string =>
  value.normalize("NFC").toLowerCase();

const markerMap = (markdown: string): MarkerIndex => {
  const found = new Map<number, Marker>();
  const ids = new Set<string>();
  for (const match of markdown.matchAll(markerPattern)) {
    const id = match[1]!;
    if (!isBlockId(id) || ids.has(id))
      throw new AgentDocxError(
        "REFERENCE_INVALID",
        `Duplicate block marker: ${id}`,
      );
    ids.add(id);
    const offset = match.index ?? 0;
    found.set(offset, { offset, end: offset + match[0].length, id });
  }
  return {
    byOffset: found,
    sorted: [...found.values()].sort(
      (left, right) => left.offset - right.offset,
    ),
  };
};

const markdownForParser = (markdown: string): string =>
  markdown.replace(markerPattern, (match) => {
    const newline = match.endsWith("\r\n")
      ? "\r\n"
      : match.endsWith("\n")
        ? "\n"
        : "";
    return `${" ".repeat(match.length - newline.length)}${newline}`;
  });

const markerBefore = (
  markers: MarkerIndex,
  markdown: string,
  offset: number,
): Marker | undefined => {
  let low = 0;
  let high = markers.sorted.length;
  while (low < high) {
    const middle = (low + high) >>> 1;
    if (markers.sorted[middle]!.end <= offset) low = middle + 1;
    else high = middle;
  }
  const marker = markers.sorted[low - 1];
  return marker && /^[ \t]*$/.test(markdown.slice(marker.end, offset))
    ? marker
    : undefined;
};

const appendInline = (
  target: InlineResult,
  text: string,
  node: MarkdownNode,
  state: Pick<InlineRun, "bold" | "italic" | "strikethrough" | "literal">,
  extra: Omit<
    Partial<InlineRun>,
    keyof typeof state | "text" | "hardBreakAfter"
  > = {},
) => {
  if (!text) return;
  const position = positionOf(node);
  const start = target.text.length;
  target.text += text;
  target.runs.push({ ...state, text, hardBreakAfter: false, ...extra });
  const nodeSource = sourceTextOf(node, inlineSources.get(target) ?? "");
  const exact = nodeSource === text;
  target.segments.push({
    normalizedStart: start,
    normalizedEnd: target.text.length,
    sourceStartOffset: position.start.offset,
    position,
    precision: exact ? "exact" : "node",
  });
};

const inlineSources = new WeakMap<InlineResult, string>();

const mergeInline = (target: InlineResult, nested: InlineResult) => {
  const offset = target.text.length;
  target.runs.push(...nested.runs);
  target.text += nested.text;
  for (const segment of nested.segments) {
    target.segments.push({
      ...segment,
      normalizedStart: segment.normalizedStart + offset,
      normalizedEnd: segment.normalizedEnd + offset,
    });
  }
};

const permittedLink = (url: string): boolean => {
  try {
    const parsed = new URL(url);
    return (
      (parsed.protocol === "http:" ||
        parsed.protocol === "https:" ||
        parsed.protocol === "mailto:") &&
      parsed.hash.length === 0
    );
  } catch {
    return false;
  }
};

const requireAttributes = (
  node: MarkdownNode,
  required: readonly string[],
  optional: readonly string[] = [],
): Record<string, string> => {
  const attributes = node.attributes ?? {};
  const allowed = new Set([...required, ...optional]);
  for (const key of Object.keys(attributes))
    if (!allowed.has(key))
      errorAt(
        "UNSUPPORTED_MARKDOWN",
        `Unknown directive attribute: ${key}`,
        node,
      );
  const result: Record<string, string> = {};
  for (const key of required) {
    const value = attributes[key];
    if (!value)
      errorAt(
        "UNSUPPORTED_MARKDOWN",
        `Missing directive attribute: ${key}`,
        node,
      );
    result[key] = value!;
  }
  for (const key of optional) {
    const value = attributes[key];
    if (value) result[key] = value;
  }
  return result;
};

const inline = (
  nodes: readonly MarkdownNode[],
  markdown: string,
  state: Pick<InlineRun, "bold" | "italic" | "strikethrough" | "literal"> = {
    bold: false,
    italic: false,
    strikethrough: false,
    literal: false,
  },
): InlineResult => {
  const result: InlineResult = { runs: [], text: "", segments: [] };
  inlineSources.set(result, markdown);
  for (const node of nodes) {
    switch (node.type) {
      case "text": {
        const value = node.value ?? "";
        const position = node.position;
        const sourceStart = position?.start.offset;
        const sourceEnd = position?.end.offset;
        if (
          value === "" ||
          sourceStart === undefined ||
          sourceEnd === undefined
        ) {
          appendInline(result, value, node, state);
          break;
        }
        // remark only emits footnoteReference nodes for references with a
        // matching definition, so missing-definition references arrive as
        // plain text. Scan the SOURCE slice (which preserves backslash
        // escapes) so escaped `\[^x]` stays literal while real references
        // are recognized and validated against the definitions.
        const source = markdown.slice(sourceStart, sourceEnd);
        const footnotePattern = /(?<!\\)\[\^([^\]]+)\]/g;
        let match: RegExpExecArray | null;
        let valueCursor = 0;
        while ((match = footnotePattern.exec(source))) {
          const remaining = value.slice(valueCursor);
          const at = remaining.indexOf(match[0]);
          if (at < 0) break;
          appendInline(result, remaining.slice(0, at), node, state);
          const label = match[0].slice(2, -1);
          appendInline(result, "⁎", node, state, {
            footnoteId: normalizeFootnoteLabel(label),
          });
          valueCursor += at + match[0].length;
        }
        appendInline(result, value.slice(valueCursor), node, state);
        break;
      }
      case "strong":
        mergeInline(
          result,
          inline(node.children ?? [], markdown, { ...state, bold: true }),
        );
        break;
      case "emphasis":
        mergeInline(
          result,
          inline(node.children ?? [], markdown, { ...state, italic: true }),
        );
        break;
      case "delete":
        mergeInline(
          result,
          inline(node.children ?? [], markdown, {
            ...state,
            strikethrough: true,
          }),
        );
        break;
      case "inlineCode":
        appendInline(result, node.value ?? "", node, {
          ...state,
          literal: true,
        });
        break;
      case "break": {
        const last = result.runs.at(-1);
        if (last) {
          last.hardBreakAfter = true;
          const position = positionOf(node);
          const normalizedStart = result.text.length;
          result.text += "\n";
          result.segments.push({
            normalizedStart,
            normalizedEnd: normalizedStart + 1,
            sourceStartOffset: position.start.offset,
            position,
            precision: "node",
          });
        } else appendInline(result, "\n", node, state);
        break;
      }
      case "link": {
        const url = node.url ?? "";
        if (!permittedLink(url))
          errorAt("REFERENCE_INVALID", `Unsupported link target: ${url}`, node);
        const nested = inline(node.children ?? [], markdown, state);
        for (const run of nested.runs)
          run.link = {
            target: url,
            ...(node.title ? { title: node.title } : {}),
          };
        mergeInline(result, nested);
        break;
      }
      case "footnoteReference": {
        const id = node.identifier
          ? normalizeFootnoteLabel(node.identifier)
          : undefined;
        if (!id)
          errorAt("REFERENCE_INVALID", "Missing footnote identifier", node);
        appendInline(result, "⁎", node, state, { footnoteId: id! });
        break;
      }
      case "textDirective": {
        if (node.name === "ref") {
          const attributes = requireAttributes(node, ["target"]);
          const target = attributes.target as BlockId;
          if (!isBlockId(target))
            errorAt(
              "REFERENCE_INVALID",
              `Invalid reference target: ${target}`,
              node,
            );
          const nested = inline(node.children ?? [], markdown, state);
          for (const run of nested.runs) run.referenceTarget = target;
          mergeInline(result, nested);
          break;
        }
        if (node.name === "authority") {
          const attributes = requireAttributes(node, [
            "id",
            "category",
            "short",
          ]);
          const category = attributes.category!;
          if (
            !(
              ["cases", "statutes", "rules", "other"] as readonly string[]
            ).includes(category)
          )
            errorAt(
              "REFERENCE_INVALID",
              `Invalid authority category: ${category}`,
              node,
            );
          const nested = inline(node.children ?? [], markdown, state);
          for (const run of nested.runs)
            run.authority = {
              id: attributes.id!,
              category: category as NonNullable<
                InlineRun["authority"]
              >["category"],
              short: attributes.short!,
            };
          mergeInline(result, nested);
          break;
        }
        errorAt(
          "UNSUPPORTED_MARKDOWN",
          `Unsupported directive: ${node.name ?? ""}`,
          node,
        );
        break;
      }
      case "image":
      case "html":
        errorAt(
          "UNSUPPORTED_MARKDOWN",
          `Unsupported Markdown node: ${node.type}`,
          node,
        );
        break;
      default:
        errorAt(
          "UNSUPPORTED_MARKDOWN",
          `Unsupported Markdown node: ${node.type}`,
          node,
        );
    }
  }
  return result;
};

const collectRefs = (runs: readonly InlineRun[]): string[] =>
  runs.flatMap((run) => (run.footnoteId ? [run.footnoteId] : []));

const validateRelativeAssetPath = (
  source: string,
  node: MarkdownNode,
): string => {
  if (!isSafeRelativePath(source))
    errorAt("REFERENCE_INVALID", `Invalid asset source: ${source}`, node);
  return source;
};

const pngDimensions = (
  bytes: Uint8Array,
): { width: number; height: number } | null => {
  if (
    bytes.byteLength < 24 ||
    bytes[0] !== 0x89 ||
    bytes[1] !== 0x50 ||
    bytes[2] !== 0x4e ||
    bytes[3] !== 0x47 ||
    bytes[4] !== 0x0d ||
    bytes[5] !== 0x0a ||
    bytes[6] !== 0x1a ||
    bytes[7] !== 0x0a
  )
    return null;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  return { width: view.getUint32(16), height: view.getUint32(20) };
};

const jpegDimensions = (
  bytes: Uint8Array,
): { width: number; height: number } | null => {
  if (bytes.byteLength < 4 || bytes[0] !== 0xff || bytes[1] !== 0xd8)
    return null;
  let offset = 2;
  while (offset + 9 < bytes.byteLength) {
    if (bytes[offset] !== 0xff) return null;
    const marker = bytes[offset + 1]!;
    offset += 2;
    if (marker === 0xd8 || marker === 0xd9) continue;
    if (offset + 2 > bytes.byteLength) return null;
    const length = (bytes[offset]! << 8) | bytes[offset + 1]!;
    if (length < 2 || offset + length > bytes.byteLength) return null;
    if (
      (marker >= 0xc0 && marker <= 0xc3) ||
      (marker >= 0xc5 && marker <= 0xc7) ||
      (marker >= 0xc9 && marker <= 0xcb) ||
      (marker >= 0xcd && marker <= 0xcf)
    ) {
      return {
        height: (bytes[offset + 3]! << 8) | bytes[offset + 4]!,
        width: (bytes[offset + 5]! << 8) | bytes[offset + 6]!,
      };
    }
    offset += length;
  }
  return null;
};

const assetFor = (
  context: ParseContext,
  source: string,
  node: MarkdownNode,
  image: boolean,
): LegalAssetInput => {
  const normalized = validateRelativeAssetPath(source, node);
  const asset = context.assets.get(normalized);
  if (!asset)
    return errorAt("REFERENCE_INVALID", `Missing asset: ${normalized}`, node);
  if (image && !["image/png", "image/jpeg"].includes(asset.mediaType))
    return errorAt(
      "REFERENCE_INVALID",
      `Image must be PNG or JPEG: ${normalized}`,
      node,
    );
  const dimensions =
    asset.mediaType === "image/png"
      ? pngDimensions(asset.bytes)
      : asset.mediaType === "image/jpeg"
        ? jpegDimensions(asset.bytes)
        : null;
  if (image && !dimensions)
    return errorAt(
      "REFERENCE_INVALID",
      `Invalid image bytes: ${normalized}`,
      node,
    );
  return asset;
};

const baseFor = (
  node: MarkdownNode,
  kind: string,
  context: ParseContext,
): {
  id: BlockId;
  position: SourcePosition;
  sourceText: string;
  segments: [];
} => {
  const position = positionOf(node);
  const sourceText = sourceTextOf(node, context.markdown);
  const marker = markerBefore(
    context.markers,
    context.markdown,
    position.start.offset,
  );
  let id: BlockId;
  if (marker) {
    id = marker.id;
    context.usedMarkers.add(marker.offset);
  } else {
    id = deterministicBlockId(
      context.options.documentId,
      kind,
      position.start.offset,
      sourceText,
    );
    context.missingMarkers.push({ offset: position.start.offset, id });
    if (context.options.requireMarkers)
      errorAt("REFERENCE_INVALID", `Missing block marker before ${kind}`, node);
  }
  return { id, position, sourceText, segments: [] };
};

const paragraphFor = (
  node: MarkdownNode,
  context: ParseContext,
): InlineParagraph => {
  const normalized = inline(node.children ?? [], context.markdown);
  return {
    position: positionOf(node),
    sourceText: sourceTextOf(node, context.markdown),
    segments: normalized.segments,
    runs: normalized.runs,
    footnoteRefs: collectRefs(normalized.runs),
  };
};

const listFor = (
  node: MarkdownNode,
  context: ParseContext,
  depth: number,
): LegalListBlock => {
  if (depth > MAX_NESTING_DEPTH)
    errorAt(
      "UNSUPPORTED_MARKDOWN",
      `Markdown nesting exceeds ${MAX_NESTING_DEPTH} levels`,
      node,
    );
  const base = baseFor(node, "list", context);
  const items: LegalListItem[] = [];
  for (const item of node.children ?? []) {
    if (item.type !== "listItem")
      errorAt("UNSUPPORTED_MARKDOWN", "Invalid list item", item);
    const paragraphs: InlineParagraph[] = [];
    const children: LegalListBlock[] = [];
    for (const child of item.children ?? []) {
      if (child.type === "paragraph")
        paragraphs.push(paragraphFor(child, context));
      else if (child.type === "list")
        children.push(listFor(child, context, depth + 1));
      else if (
        child.type === "html" &&
        markerBefore(
          context.markers,
          context.markdown,
          positionOf(child).start.offset,
        )
      ) {
        continue;
      } else
        errorAt(
          "UNSUPPORTED_MARKDOWN",
          `Unsupported list content: ${child.type}`,
          child,
        );
    }
    if (paragraphs.length === 0)
      errorAt("UNSUPPORTED_MARKDOWN", "List item requires a paragraph", item);
    items.push({
      position: positionOf(item),
      sourceText: sourceTextOf(item, context.markdown),
      segments: paragraphs.flatMap((paragraph) => paragraph.segments),
      paragraphs,
      children,
    });
  }
  return {
    ...base,
    kind: "list",
    ordered: node.ordered === true,
    start: node.ordered === true ? (node.start ?? 1) : null,
    depth,
    items,
  };
};

const tableFor = (
  node: MarkdownNode,
  context: ParseContext,
): Extract<LegalBlock, { kind: "table" }> => {
  const base = baseFor(node, "table", context);
  const sourceRows = node.children ?? [];
  const width = sourceRows.reduce(
    (maximum, row) => Math.max(maximum, row.children?.length ?? 0),
    0,
  );
  const emptyParagraph = (row: MarkdownNode): InlineParagraph => ({
    position: positionOf(row),
    sourceText: "",
    segments: [],
    runs: [],
    footnoteRefs: [],
  });
  const rows = sourceRows.map((row) => {
    if (row.type !== "tableRow")
      errorAt("UNSUPPORTED_MARKDOWN", "Invalid table row", row);
    return Array.from({ length: width }, (_, columnIndex): LegalTableCell => {
      const cell = row.children?.[columnIndex];
      if (!cell)
        return {
          paragraphs: [emptyParagraph(row)],
          footnoteRefs: [],
          verticalAlign: "top",
        };
      if (cell.type !== "tableCell")
        errorAt("UNSUPPORTED_MARKDOWN", "Invalid table cell", cell);
      const paragraph = paragraphFor(cell, context);
      return {
        paragraphs: [paragraph],
        footnoteRefs: paragraph.footnoteRefs,
        verticalAlign: "top",
      };
    });
  });
  return {
    ...base,
    kind: "table",
    rows,
    align: Array.from(
      { length: width },
      (_, columnIndex) => node.align?.[columnIndex] ?? null,
    ),
  };
};

const leafDirectiveFor = (
  node: MarkdownNode,
  context: ParseContext,
): LegalBlock => {
  const base = baseFor(node, node.name ?? "directive", context);
  const name = node.name;
  if (
    name === "caption" ||
    name === "toc" ||
    name === "toa" ||
    name === "pagebreak"
  ) {
    requireAttributes(node, []);
    return {
      ...base,
      kind: name as "caption" | "toc" | "toa" | "pagebreak",
    };
  }
  if (name === "signature") {
    const attributes = requireAttributes(node, ["counsel"]);
    return {
      ...base,
      kind: "signature",
      counselId: attributes.counsel!,
    };
  }
  if (name === "certificate") {
    const attributes = requireAttributes(node, ["id"]);
    return {
      ...base,
      kind: "certificate",
      certificateId: attributes.id!,
    };
  }
  if (name === "sectionbreak") {
    const attributes = requireAttributes(
      node,
      ["kind"],
      ["pageNumberFormat", "pageNumberStart"],
    );
    const breakKind = attributes.kind!;
    if (!["next-page", "continuous"].includes(breakKind))
      return errorAt(
        "REFERENCE_INVALID",
        `Invalid section break kind: ${breakKind}`,
        node,
      );
    const format = attributes.pageNumberFormat;
    const start = attributes.pageNumberStart;
    if ((format === undefined) !== (start === undefined))
      errorAt(
        "REFERENCE_INVALID",
        "Section page number attributes are paired",
        node,
      );
    if (format && !["decimal", "lower-roman", "upper-roman"].includes(format))
      errorAt(
        "REFERENCE_INVALID",
        `Invalid page number format: ${format}`,
        node,
      );
    if (
      start &&
      (!/^[1-9][0-9]*$/.test(start) || !Number.isSafeInteger(Number(start)))
    )
      errorAt("REFERENCE_INVALID", `Invalid page number start: ${start}`, node);
    return {
      ...base,
      kind: "sectionbreak",
      breakKind: breakKind as "next-page" | "continuous",
      ...(format && start
        ? {
            pageNumber: {
              format: format as "decimal" | "lower-roman" | "upper-roman",
              start: Number(start),
            },
          }
        : {}),
    };
  }
  if (name === "image") {
    const attributes = requireAttributes(node, [
      "source",
      "alt",
      "widthTwips",
      "heightTwips",
    ]);
    const width = Number(attributes.widthTwips);
    const height = Number(attributes.heightTwips);
    if (
      !Number.isInteger(width) ||
      !Number.isInteger(height) ||
      width <= 0 ||
      height <= 0
    )
      return errorAt(
        "REFERENCE_INVALID",
        "Image dimensions must be positive twips",
        node,
      );
    const source = validateRelativeAssetPath(attributes.source!, node);
    assetFor(context, source, node, true);
    return {
      ...base,
      kind: "image",
      source,
      alt: attributes.alt!,
      widthTwips: width,
      heightTwips: height,
    };
  }
  return errorAt(
    "UNSUPPORTED_MARKDOWN",
    `Unsupported directive: ${name ?? ""}`,
    node,
  );
};

const containerDirectiveFor = (
  node: MarkdownNode,
  context: ParseContext,
  depth: number,
): LegalBlock => {
  const name = node.name;
  const base = baseFor(node, name ?? "directive", context);
  if (name === "numbered") {
    const attributes = requireAttributes(node, ["sequence", "level"]);
    const sequence = attributes.sequence!;
    const level = attributes.level!;
    if (!/^[1-4]$/.test(level))
      return errorAt(
        "REFERENCE_INVALID",
        `Invalid numbered level: ${level}`,
        node,
      );
    const children = node.children ?? [];
    if (children.length !== 1 || children[0]!.type !== "paragraph")
      errorAt("UNSUPPORTED_MARKDOWN", "numbered requires one paragraph", node);
    const normalized = inline(children[0]!.children ?? [], context.markdown);
    return {
      ...base,
      kind: "numbered-paragraph",
      sequence,
      level: Number(level) as 1 | 2 | 3 | 4,
      runs: normalized.runs,
      segments: normalized.segments,
      footnoteRefs: collectRefs(normalized.runs),
    };
  }
  if (name === "exhibit") {
    const attributes = requireAttributes(node, ["id", "label", "source"]);
    const source = validateRelativeAssetPath(attributes.source!, node);
    assetFor(context, source, node, false);
    return {
      ...base,
      kind: "exhibit",
      exhibitId: attributes.id!,
      label: attributes.label!,
      source,
      blocks: parseBlocks(node.children ?? [], context, depth + 1),
    };
  }
  if (name === "length-exclusion") {
    const attributes = requireAttributes(node, ["kind"], ["citation"]);
    const kind = attributes.kind as Extract<
      LegalBlock,
      { kind: "length-exclusion" }
    >["exclusionKind"];
    if (
      ![
        "disclosure-statement",
        "oral-argument-statement",
        "statutory-addendum",
        "proof-of-service",
        "local-rule",
      ].includes(kind)
    )
      return errorAt(
        "REFERENCE_INVALID",
        `Invalid length exclusion: ${attributes.kind}`,
        node,
      );
    if ((kind === "local-rule") !== (attributes.citation !== undefined))
      return errorAt(
        "REFERENCE_INVALID",
        "local-rule requires citation and other exclusions forbid it",
        node,
      );
    return {
      ...base,
      kind: "length-exclusion",
      exclusionKind: kind,
      ...(attributes.citation ? { citation: attributes.citation } : {}),
      blocks: parseBlocks(node.children ?? [], context, depth + 1),
    };
  }
  return errorAt(
    "UNSUPPORTED_MARKDOWN",
    `Unsupported directive: ${name ?? ""}`,
    node,
  );
};

const parseBlock = (
  node: MarkdownNode,
  context: ParseContext,
  depth: number,
): LegalBlock | null => {
  if (depth > MAX_NESTING_DEPTH)
    errorAt(
      "UNSUPPORTED_MARKDOWN",
      `Markdown nesting exceeds ${MAX_NESTING_DEPTH} levels`,
      node,
    );
  switch (node.type) {
    case "html": {
      const marker = markerHtmlPattern.exec(node.value ?? "");
      if (marker && isBlockId(marker[1]!)) return null;
      return errorAt(
        "UNSUPPORTED_MARKDOWN",
        "Arbitrary HTML is unsupported",
        node,
      );
    }
    case "paragraph": {
      const base = baseFor(node, "paragraph", context);
      const normalized = inline(node.children ?? [], context.markdown);
      return {
        ...base,
        kind: "paragraph",
        runs: normalized.runs,
        segments: normalized.segments,
        footnoteRefs: collectRefs(normalized.runs),
      };
    }
    case "heading": {
      const level = node.depth;
      if (!level || level < 1 || level > 6)
        errorAt("UNSUPPORTED_MARKDOWN", "Invalid heading level", node);
      const base = baseFor(node, "heading", context);
      const normalized = inline(node.children ?? [], context.markdown);
      return {
        ...base,
        kind: "heading",
        level: level as 1 | 2 | 3 | 4 | 5 | 6,
        runs: normalized.runs,
        segments: normalized.segments,
        footnoteRefs: collectRefs(normalized.runs),
      };
    }
    case "blockquote": {
      const children = node.children ?? [];
      if (children.length !== 1 || children[0]!.type !== "paragraph")
        errorAt(
          "UNSUPPORTED_MARKDOWN",
          "Blockquotes require one paragraph",
          node,
        );
      const base = baseFor(node, "blockquote", context);
      const normalized = inline(children[0]!.children ?? [], context.markdown);
      return {
        ...base,
        kind: "blockquote",
        depth,
        runs: normalized.runs,
        segments: normalized.segments,
        footnoteRefs: collectRefs(normalized.runs),
      };
    }
    case "list":
      return listFor(node, context, depth + 1);
    case "table":
      return tableFor(node, context);
    case "thematicBreak": {
      const base = baseFor(node, "thematic-break", context);
      return { ...base, kind: "thematic-break" };
    }
    case "leafDirective":
      return leafDirectiveFor(node, context);
    case "containerDirective":
      return containerDirectiveFor(node, context, depth);
    case "code":
    case "yaml":
    case "math":
    case "image":
      return errorAt(
        "UNSUPPORTED_MARKDOWN",
        `Unsupported Markdown node: ${node.type}`,
        node,
      );
    default:
      return errorAt(
        "UNSUPPORTED_MARKDOWN",
        `Unsupported Markdown node: ${node.type}`,
        node,
      );
  }
};

const parseBlocks = (
  nodes: readonly MarkdownNode[],
  context: ParseContext,
  depth: number,
): LegalBlock[] => {
  if (depth > MAX_NESTING_DEPTH) {
    const node = nodes[0];
    if (node)
      errorAt(
        "UNSUPPORTED_MARKDOWN",
        `Markdown nesting exceeds ${MAX_NESTING_DEPTH} levels`,
        node,
      );
    throw new AgentDocxError(
      "UNSUPPORTED_MARKDOWN",
      `Markdown nesting exceeds ${MAX_NESTING_DEPTH} levels`,
    );
  }
  const blocks: LegalBlock[] = [];
  for (const node of nodes) {
    const block = parseBlock(node, context, depth);
    if (block) blocks.push(block);
  }
  return blocks;
};

const parseFootnotes = (
  nodes: readonly MarkdownNode[],
  context: ParseContext,
): FootnoteDefinition[] => {
  const footnotes: FootnoteDefinition[] = [];
  const labels = new Set<string>();
  for (const node of nodes) {
    if (node.type !== "footnoteDefinition") continue;
    const label = node.identifier
      ? normalizeFootnoteLabel(node.identifier)
      : undefined;
    if (!label || labels.has(label))
      return errorAt(
        "REFERENCE_INVALID",
        `Missing or duplicate footnote: ${label ?? ""}`,
        node,
      );
    labels.add(label);
    const base = baseFor(node, "footnote", context);
    const paragraphs = (node.children ?? []).map((child) => {
      if (child.type !== "paragraph")
        errorAt(
          "UNSUPPORTED_MARKDOWN",
          "Footnotes only support paragraphs",
          child,
        );
      return paragraphFor(child, context);
    });
    if (paragraphs.length === 0)
      errorAt("UNSUPPORTED_MARKDOWN", "Footnote requires a paragraph", node);
    footnotes.push({ ...base, kind: "footnote", label, paragraphs });
  }
  return footnotes;
};

const allBlocks = (blocks: readonly LegalBlock[]): LegalBlock[] => {
  const result: LegalBlock[] = [];
  const pending = blocks.map((block) => ({ block, depth: 0 })).reverse();
  while (pending.length > 0) {
    const current = pending.pop()!;
    if (current.depth > MAX_NESTING_DEPTH)
      throw new AgentDocxError(
        "UNSUPPORTED_MARKDOWN",
        `Markdown nesting exceeds ${MAX_NESTING_DEPTH} levels`,
      );
    result.push(current.block);
    if (
      current.block.kind === "exhibit" ||
      current.block.kind === "length-exclusion"
    )
      for (let index = current.block.blocks.length - 1; index >= 0; index--)
        pending.push({
          block: current.block.blocks[index]!,
          depth: current.depth + 1,
        });
    if (current.block.kind === "list")
      for (
        let itemIndex = current.block.items.length - 1;
        itemIndex >= 0;
        itemIndex--
      )
        for (
          let childIndex = current.block.items[itemIndex]!.children.length - 1;
          childIndex >= 0;
          childIndex--
        )
          pending.push({
            block: current.block.items[itemIndex]!.children[childIndex]!,
            depth: current.depth + 1,
          });
  }
  return result;
};

const footnoteRefsForBlock = (root: LegalBlock): readonly string[] => {
  const refs: string[] = [];
  const pending: Array<{ block: LegalBlock; depth: number }> = [
    { block: root, depth: 0 },
  ];
  while (pending.length > 0) {
    const current = pending.pop()!;
    if (current.depth > MAX_NESTING_DEPTH)
      throw new AgentDocxError(
        "UNSUPPORTED_MARKDOWN",
        `Markdown nesting exceeds ${MAX_NESTING_DEPTH} levels`,
      );
    const block = current.block;
    if ("runs" in block) refs.push(...collectRefs(block.runs));
    if (block.kind === "list") {
      for (const item of block.items) {
        for (const paragraph of item.paragraphs)
          refs.push(...paragraph.footnoteRefs);
        for (const child of item.children)
          pending.push({ block: child, depth: current.depth + 1 });
      }
    } else if (block.kind === "table") {
      for (const row of block.rows)
        for (const cell of row) refs.push(...cell.footnoteRefs);
    } else if (block.kind === "exhibit" || block.kind === "length-exclusion") {
      for (const child of block.blocks)
        pending.push({ block: child, depth: current.depth + 1 });
    }
  }
  return refs;
};

export function parseLegalMarkdown(
  markdown: string,
  options: ParseLegalMarkdownOptions,
): ParsedLegalMarkdown {
  if (typeof markdown !== "string")
    throw new AgentDocxError("INVALID_ARGUMENT", "markdown must be a string");
  if (!isDocumentId(options.documentId))
    throw new AgentDocxError("INVALID_ARGUMENT", "documentId is invalid");
  assertXml10Legal(markdown);
  const root = unified()
    .use(remarkParse)
    .use(remarkGfm)
    .use(remarkDirective)
    .parse(markdownForParser(markdown)) as unknown as MarkdownNode;
  const suppliedAssets = new Map(Object.entries(options.assets ?? {}));
  const context: ParseContext = {
    markdown,
    options,
    markers: markerMap(markdown),
    usedMarkers: new Set<number>(),
    missingMarkers: [],
    assets: suppliedAssets,
  };
  const footnotes = parseFootnotes(root.children ?? [], context);
  const blocks = parseBlocks(
    (root.children ?? []).filter((node) => node.type !== "footnoteDefinition"),
    context,
    0,
  );
  for (const marker of context.markers.byOffset.values())
    if (!context.usedMarkers.has(marker.offset))
      throw new AgentDocxError(
        "REFERENCE_INVALID",
        `Orphan block marker: ${marker.id}`,
      );

  const flattenedBlocks = allBlocks(blocks);
  const labels = new Set(footnotes.map((footnote) => footnote.label));
  const ids = new Set<string>();
  for (const block of [...flattenedBlocks, ...footnotes]) {
    if (ids.has(block.id))
      throw new AgentDocxError(
        "REFERENCE_INVALID",
        `Duplicate block ID: ${block.id}`,
      );
    ids.add(block.id);
  }
  for (const block of flattenedBlocks)
    for (const footnote of footnoteRefsForBlock(block))
      if (!labels.has(footnote))
        throw new AgentDocxError(
          "REFERENCE_INVALID",
          `Missing footnote definition: ${footnote}`,
        );
  for (const footnote of footnotes)
    for (const paragraph of footnote.paragraphs)
      for (const ref of paragraph.footnoteRefs)
        if (!labels.has(ref))
          throw new AgentDocxError(
            "REFERENCE_INVALID",
            `Missing footnote definition: ${ref}`,
          );
  const referencedAssets = new Set<string>();
  for (const block of flattenedBlocks) {
    if (block.kind === "image" || block.kind === "exhibit")
      referencedAssets.add(block.source);
  }
  if (options.exactAssets)
    for (const key of suppliedAssets.keys())
      if (!referencedAssets.has(key))
        throw new AgentDocxError(
          "REFERENCE_INVALID",
          `Unexpected asset: ${key}`,
        );
  const assets: Record<
    string,
    { sha256: `sha256:${string}`; mediaType: string; bytes: number }
  > = {};
  for (const key of [...referencedAssets].sort()) {
    const asset = suppliedAssets.get(key);
    if (!asset)
      throw new AgentDocxError("REFERENCE_INVALID", `Missing asset: ${key}`);
    assets[key] = {
      sha256: sha256(asset.bytes),
      mediaType: asset.mediaType,
      bytes: asset.bytes.byteLength,
    };
  }

  return {
    document: {
      schemaVersion: 1,
      projectId: options.projectId ?? "standalone",
      documentId: options.documentId,
      metadata: options.metadata ?? emptyLitigationMetadata(),
      chrome: options.chrome ?? {},
      blocks,
      footnotes,
      annotations: options.annotations ?? [],
      assets,
      source: { text: markdown, sha256: sha256(markdown) },
    },
    missingMarkers: context.missingMarkers,
  };
}

export function insertMissingBlockMarkers(
  markdown: string,
  options: ParseLegalMarkdownOptions,
): string {
  const parsed = parseLegalMarkdown(markdown, {
    ...options,
    requireMarkers: false,
  });
  if (parsed.missingMarkers.length === 0) return markdown;
  let result = markdown;
  for (const marker of [...parsed.missingMarkers].sort(
    (left, right) => right.offset - left.offset,
  )) {
    const lineStart = markdown.lastIndexOf("\n", marker.offset - 1) + 1;
    const indentation = markdown.slice(lineStart, marker.offset);
    result = `${result.slice(0, lineStart)}${indentation}<!-- agent-docx:block id="${marker.id}" -->\n${result.slice(lineStart)}`;
  }
  return result;
}
