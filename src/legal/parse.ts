import { createHash } from "node:crypto";
import remarkDirective from "remark-directive";
import remarkGfm from "remark-gfm";
import remarkParse from "remark-parse";
import { unified } from "unified";
import type { NormalizedSourceSegment } from "../markdown.js";
import { AgentDocxError, type SourcePosition } from "../types.js";
import {
  type BlockId,
  type DocumentChrome,
  type FootnoteDefinition,
  type InlineParagraph,
  type InlineRun,
  type LegalBlock,
  type LegalDocument,
  type LegalDocumentSpecification,
  type LegalListBlock,
  type LegalListItem,
  type LegalTableCell,
  type LitigationMetadata,
  emptyLitigationMetadata,
  isBlockId,
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
type InlineResult = {
  runs: InlineRun[];
  text: string;
  segments: NormalizedSourceSegment[];
};
type ParseContext = {
  markdown: string;
  options: ParseLegalMarkdownOptions;
  markers: ReadonlyMap<number, Marker>;
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
  /^[ \t]*<!--[ \t]*agent-docx:block[ \t]+id="(b_[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12})"[ \t]*-->\r?\n/gm;

const markerHtmlPattern =
  /^[ \t]*<!--[ \t]*agent-docx:block[ \t]+id="(b_[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12})"[ \t]*-->$/;

const markerMap = (markdown: string): ReadonlyMap<number, Marker> => {
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
  return found;
};
const markdownForParser = (markdown: string): string =>
  markdown.replace(markerPattern, (match) => {
    const newline = match.endsWith("\r\n") ? "\r\n" : "\n";
    return `${" ".repeat(match.length - newline.length)}${newline}`;
  });

const markerBefore = (
  markers: ReadonlyMap<number, Marker>,
  markdown: string,
  offset: number,
): Marker | undefined => {
  for (const marker of markers.values())
    if (
      marker.end <= offset &&
      /^[ \t]*$/.test(markdown.slice(marker.end, offset))
    )
      return marker;
  return undefined;
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

const requiredAttribute = (
  attributes: Record<string, string>,
  key: string,
  node: MarkdownNode,
): string => {
  const value = attributes[key];
  if (!value)
    return errorAt(
      "UNSUPPORTED_MARKDOWN",
      `Missing directive attribute: ${key}`,
      node,
    );
  return value;
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
      case "text":
        appendInline(result, node.value ?? "", node, state);
        break;
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
        if (last) last.hardBreakAfter = true;
        else appendInline(result, "\n", node, state);
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
        const id = node.identifier?.toLowerCase();
        if (!id)
          errorAt("REFERENCE_INVALID", "Missing footnote identifier", node);
        appendInline(result, "⁎", node, state, { footnoteId: id });
        break;
      }
      case "textDirective": {
        if (node.name === "ref") {
          const attributes = requireAttributes(node, ["target"]);
          const target = requiredAttribute(
            attributes,
            "target",
            node,
          ) as BlockId;
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
          const category = requiredAttribute(attributes, "category", node);
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
              id: requiredAttribute(attributes, "id", node),
              category: category as NonNullable<
                InlineRun["authority"]
              >["category"],
              short: requiredAttribute(attributes, "short", node),
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
  if (
    source.length === 0 ||
    source.startsWith("/") ||
    source.includes("\\") ||
    source
      .split("/")
      .some((part) => part === "" || part === "." || part === "..")
  )
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
  };
};

const listFor = (
  node: MarkdownNode,
  context: ParseContext,
  depth: number,
): LegalListBlock => {
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
  const rows = (node.children ?? []).map((row) => {
    if (row.type !== "tableRow")
      errorAt("UNSUPPORTED_MARKDOWN", "Invalid table row", row);
    return (row.children ?? []).map((cell): LegalTableCell => {
      if (cell.type !== "tableCell")
        errorAt("UNSUPPORTED_MARKDOWN", "Invalid table cell", cell);
      return {
        paragraphs: [paragraphFor(cell, context)],
        verticalAlign: "top",
      };
    });
  });
  return {
    ...base,
    kind: "table",
    rows,
    align: node.align ?? [],
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
      counselId: requiredAttribute(attributes, "counsel", node),
    };
  }
  if (name === "certificate") {
    const attributes = requireAttributes(node, ["id"]);
    return {
      ...base,
      kind: "certificate",
      certificateId: requiredAttribute(attributes, "id", node),
    };
  }
  if (name === "sectionbreak") {
    const attributes = requireAttributes(
      node,
      ["kind"],
      ["pageNumberFormat", "pageNumberStart"],
    );
    const breakKind = requiredAttribute(attributes, "kind", node);
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
    const width = Number(requiredAttribute(attributes, "widthTwips", node));
    const height = Number(requiredAttribute(attributes, "heightTwips", node));
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
    const source = validateRelativeAssetPath(
      requiredAttribute(attributes, "source", node),
      node,
    );
    assetFor(context, source, node, true);
    return {
      ...base,
      kind: "image",
      source,
      alt: requiredAttribute(attributes, "alt", node),
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
    const sequence = requiredAttribute(attributes, "sequence", node);
    const level = requiredAttribute(attributes, "level", node);
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
    };
  }
  if (name === "exhibit") {
    const attributes = requireAttributes(node, ["id", "label", "source"]);
    const source = validateRelativeAssetPath(
      requiredAttribute(attributes, "source", node),
      node,
    );
    assetFor(context, source, node, false);
    return {
      ...base,
      kind: "exhibit",
      exhibitId: requiredAttribute(attributes, "id", node),
      label: requiredAttribute(attributes, "label", node),
      source,
      blocks: parseBlocks(node.children ?? [], context, depth),
    };
  }
  if (name === "length-exclusion") {
    const attributes = requireAttributes(node, ["kind"], ["citation"]);
    const kind = requiredAttribute(attributes, "kind", node) as Extract<
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
      blocks: parseBlocks(node.children ?? [], context, depth),
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
    const label = node.identifier?.toLowerCase();
    if (!label || labels.has(label))
      return errorAt(
        "REFERENCE_INVALID",
        `Missing or duplicate footnote: ${label ?? ""}`,
        node,
      );
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
    footnotes.push({ ...base, kind: "footnote", label: label!, paragraphs });
  }
  return footnotes;
};

const allBlocks = (blocks: readonly LegalBlock[]): LegalBlock[] =>
  blocks.flatMap((block) => {
    if (block.kind === "exhibit" || block.kind === "length-exclusion")
      return [block, ...allBlocks(block.blocks)];
    if (block.kind === "list")
      return [
        block,
        ...block.items.flatMap((item) =>
          item.children.flatMap((child) => allBlocks([child])),
        ),
      ];
    return [block];
  });

export function parseLegalMarkdown(
  markdown: string,
  options: ParseLegalMarkdownOptions,
): ParsedLegalMarkdown {
  if (typeof markdown !== "string")
    throw new AgentDocxError("INVALID_ARGUMENT", "markdown must be a string");
  if (!options.documentId)
    throw new AgentDocxError("INVALID_ARGUMENT", "documentId is required");
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
  for (const marker of context.markers.values())
    if (!context.usedMarkers.has(marker.offset))
      throw new AgentDocxError(
        "REFERENCE_INVALID",
        `Orphan block marker: ${marker.id}`,
      );

  const labels = new Set(footnotes.map((footnote) => footnote.label));
  const ids = new Set<string>();
  for (const block of [...allBlocks(blocks), ...footnotes]) {
    if (ids.has(block.id))
      throw new AgentDocxError(
        "REFERENCE_INVALID",
        `Duplicate block ID: ${block.id}`,
      );
    ids.add(block.id);
  }
  for (const block of allBlocks(blocks)) {
    const refs =
      block.kind === "paragraph" || block.kind === "blockquote"
        ? block.footnoteRefs
        : [];
    for (const footnote of refs)
      if (!labels.has(footnote))
        throw new AgentDocxError(
          "REFERENCE_INVALID",
          `Missing footnote definition: ${footnote}`,
        );
  }
  for (const footnote of footnotes)
    for (const paragraph of footnote.paragraphs)
      for (const ref of collectRefs(paragraph.runs))
        if (!labels.has(ref))
          throw new AgentDocxError(
            "REFERENCE_INVALID",
            `Missing footnote definition: ${ref}`,
          );

  const referencedAssets = new Set<string>();
  for (const block of allBlocks(blocks)) {
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

export function documentFromSpecification(
  markdown: string,
  specification: LegalDocumentSpecification,
): ParsedLegalMarkdown {
  return parseLegalMarkdown(markdown, {
    projectId: specification.projectId,
    documentId: specification.documentId,
    metadata: specification.metadata,
    chrome: specification.chrome,
    assets: specification.assets,
    exactAssets: true,
  });
}
