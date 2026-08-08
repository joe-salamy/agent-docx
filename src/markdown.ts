import { unified } from "unified";
import remarkParse from "remark-parse";
import remarkGfm from "remark-gfm";
import remarkFrontmatter from "remark-frontmatter";
import remarkMath from "remark-math";
import {
  AgentDocxError,
  type SectionHeading,
  type SourcePosition,
} from "./types.js";
import type { AuthorityReference, LegalBlock } from "./legal/model.js";
type Node = {
  type: string;
  value?: string;
  depth?: number;
  ordered?: boolean;
  children?: Node[];
  align?: Array<"left" | "center" | "right" | null>;
  position?: {
    start: { line: number; column: number; offset?: number };
    end: { line: number; column: number; offset?: number };
  };
  identifier?: string;
  label?: string;
  url?: string;
  title?: string | null;
};
export type NormalizedSourceSegment = {
  normalizedStart: number;
  normalizedEnd: number;
  sourceStartOffset: number;
  position: SourcePosition;
  precision: "exact" | "node";
};
export type InlineRun = {
  text: string;
  bold: boolean;
  italic: boolean;
  footnoteId?: string;
  literal?: boolean;
  strikethrough?: boolean;
  link?: { target: string; title?: string };
  authority?: AuthorityReference;
};
export type TextBlockKind =
  | "paragraph"
  | "heading"
  | "blockquote"
  | "list"
  | "footnote";
export type TextFlowBlock = {
  kind: TextBlockKind;
  runs: InlineRun[];
  normalizedText: string;
  sourceSegments: NormalizedSourceSegment[];
  position: SourcePosition;
  level?: number;
  footnoteRefs: string[];
  legalBlockId?: string;
  legalKind?: LegalBlock["kind"] | "footnote";
  listOrdered?: boolean;
  listLevel?: number;
  numberedLevel?: number;
  listStart?: number;
  image?: {
    source: string;
    alt: string;
    widthTwips: number;
    heightTwips: number;
  };
};
export type TableCell = InlineNormalization & {
  position: SourcePosition;
  alignment: "left" | "center" | "right" | null;
};
export type TableFlowBlock = {
  kind: "table";
  position: SourcePosition;
  alignments: readonly ("left" | "center" | "right" | null)[];
  rows: readonly (readonly TableCell[])[];
};
export type ThematicBreakFlowBlock = {
  kind: "thematic-break";
  position: SourcePosition;
};
export type PageBreakFlowBlock = {
  kind: "pagebreak";
  position: SourcePosition;
  sectionBreak?: {
    kind: "next-page" | "continuous";
    pageNumber?: {
      format: "decimal" | "lower-roman" | "upper-roman";
      start: number;
    };
  };
};
export type FlowBlock =
  | TextFlowBlock
  | TableFlowBlock
  | ThematicBreakFlowBlock
  | PageBreakFlowBlock;
export type FootnoteDefinition = {
  id: string;
  position: SourcePosition;
  blocks: readonly TextFlowBlock[];
  footnoteRefs: readonly string[];
};
export type NormalizedDocument = {
  blocks: FlowBlock[];
  footnotes: Map<string, FootnoteDefinition>;
  paragraphs: TextFlowBlock[];
};
export type IndexedSection = {
  index: number;
  parentIndex: number | null;
  heading: SectionHeading | null;
  position: SourcePosition | null;
  empty: boolean;
  ancestors: readonly number[];
};
export type SectionIndex = {
  sections: readonly IndexedSection[];
  deepestOwnerByBlock: ReadonlyMap<FlowBlock, number>;
};

export function indexSections(blocks: readonly FlowBlock[]): SectionIndex {
  type WorkSection = IndexedSection & {
    firstPosition: SourcePosition | null;
    lastPosition: SourcePosition | null;
  };
  const sections: WorkSection[] = [
    {
      index: 0,
      parentIndex: null,
      heading: null,
      position: null,
      empty: true,
      ancestors: [0],
      firstPosition: null,
      lastPosition: null,
    },
  ];
  const stack: number[] = [];
  const deepestOwnerByBlock = new Map<FlowBlock, number>();
  for (const block of blocks) {
    let owner = stack.at(-1) ?? 0;
    if (block.kind === "heading") {
      const level = block.level as 1 | 2 | 3 | 4 | 5 | 6;
      while (stack.length && sections[stack.at(-1)!]!.heading!.level >= level) {
        stack.pop();
      }
      const parentIndex = stack.at(-1) ?? null;
      const index = sections.length;
      const heading: SectionHeading = {
        level,
        title: block.runs
          .filter((run) => run.footnoteId === undefined)
          .map((run) => run.text)
          .join("")
          .replace(/\s+/gu, " ")
          .trim(),
        position: block.position,
      };
      sections.push({
        index,
        parentIndex,
        heading,
        position: block.position,
        empty: true,
        ancestors:
          parentIndex === null
            ? [index]
            : [...sections[parentIndex]!.ancestors, index],
        firstPosition: block.position,
        lastPosition: block.position,
      });
      stack.push(index);
      owner = index;
    }
    deepestOwnerByBlock.set(block, owner);
    const owners = sections[owner]!.ancestors;
    for (const sectionIndex of owners) {
      const section = sections[sectionIndex]!;
      section.firstPosition ??= block.position;
      section.lastPosition = block.position;
      if (block.kind !== "heading" && block.kind !== "pagebreak") {
        section.empty = false;
      }
    }
  }
  return {
    sections: sections.map(
      ({ firstPosition, lastPosition, ...section }): IndexedSection => ({
        ...section,
        position:
          firstPosition === null || lastPosition === null
            ? null
            : {
                start: firstPosition.start,
                end: lastPosition.end,
              },
      }),
    ),
    deepestOwnerByBlock,
  };
}
const pos = (node: Node): SourcePosition => {
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
const lineStartsFor = (markdown: string): number[] => {
  const starts = [0];
  for (let i = 0; i < markdown.length; i++)
    if (markdown.charCodeAt(i) === 10) starts.push(i + 1);
  return starts;
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
const assertXml10Legal = (markdown: string): void => {
  const lineStarts = lineStartsFor(markdown);
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
            start: sourcePoint(offset, lineStarts),
            end: sourcePoint(offset + width, lineStarts),
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
            start: sourcePoint(offset, lineStarts),
            end: sourcePoint(offset + width, lineStarts),
          } as unknown as Record<string, never>,
          codePoint: label,
        },
      );
    }
    offset += width;
  }
};
const sourcePoint = (
  offset: number,
  lineStarts: readonly number[],
): SourcePosition["start"] => {
  let low = 0;
  let high = lineStarts.length;
  while (low + 1 < high) {
    const middle = (low + high) >>> 1;
    if (lineStarts[middle]! <= offset) low = middle;
    else high = middle;
  }
  return {
    line: low + 1,
    column: offset - lineStarts[low]! + 1,
    offset,
  };
};
const sourceRange = (
  start: number,
  end: number,
  lineStarts: readonly number[],
): SourcePosition => ({
  start: sourcePoint(start, lineStarts),
  end: sourcePoint(end, lineStarts),
});
const normalizeFootnoteLabel = (value: string): string =>
  value.normalize("NFC").toLowerCase();

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
const MAX_NESTING_DEPTH = 100;

const unsupported = (node: Node): never => {
  throw new AgentDocxError(
    "UNSUPPORTED_MARKDOWN",
    `Unsupported Markdown node: ${node.type}`,
    { position: pos(node) as unknown as Record<string, never> },
  );
};
type InlineNormalization = {
  runs: InlineRun[];
  normalizedText: string;
  sourceSegments: NormalizedSourceSegment[];
  footnoteRefs: string[];
};

function inline(
  nodes: Node[],
  markdown: string,
  lineStarts: readonly number[],
  bold = false,
  italic = false,
  refs: string[] = [],
  strikethrough = false,
): InlineNormalization {
  const runs: InlineRun[] = [];
  const sourceSegments: NormalizedSourceSegment[] = [];
  let normalizedText = "";
  const append = (
    text: string,
    node: Node,
    run: Omit<InlineRun, "text">,
    exactSourceOffset?: number,
  ) => {
    if (!text) return;
    const normalizedStart = normalizedText.length;
    normalizedText += text;
    runs.push({ text, ...run });
    const normalizedEnd = normalizedText.length;
    const precision = exactSourceOffset === undefined ? "node" : "exact";
    const position =
      exactSourceOffset !== undefined
        ? sourceRange(
            exactSourceOffset,
            exactSourceOffset + text.length,
            lineStarts,
          )
        : pos(node);
    const segment: NormalizedSourceSegment = {
      normalizedStart,
      normalizedEnd,
      sourceStartOffset:
        exactSourceOffset ??
        node.position?.start.offset ??
        position.start.offset,
      position,
      precision,
    };
    const previous = sourceSegments.at(-1);
    if (
      previous?.precision === "exact" &&
      segment.precision === "exact" &&
      previous.normalizedEnd === segment.normalizedStart &&
      previous.sourceStartOffset +
        (previous.normalizedEnd - previous.normalizedStart) ===
        segment.sourceStartOffset
    ) {
      previous.normalizedEnd = segment.normalizedEnd;
      previous.position = sourceRange(
        previous.sourceStartOffset,
        segment.sourceStartOffset + text.length,
        lineStarts,
      );
    } else {
      sourceSegments.push(segment);
    }
  };
  const appendNested = (nested: InlineNormalization) => {
    const normalizedOffset = normalizedText.length;
    runs.push(...nested.runs);
    normalizedText += nested.normalizedText;
    for (const nestedSegment of nested.sourceSegments) {
      const segment = {
        ...nestedSegment,
        normalizedStart: nestedSegment.normalizedStart + normalizedOffset,
        normalizedEnd: nestedSegment.normalizedEnd + normalizedOffset,
      };
      const previous = sourceSegments.at(-1);
      if (
        previous?.precision === "exact" &&
        segment.precision === "exact" &&
        previous.normalizedEnd === segment.normalizedStart &&
        previous.sourceStartOffset +
          (previous.normalizedEnd - previous.normalizedStart) ===
          segment.sourceStartOffset
      ) {
        previous.normalizedEnd = segment.normalizedEnd;
        previous.position = sourceRange(
          previous.sourceStartOffset,
          segment.sourceStartOffset +
            segment.normalizedEnd -
            segment.normalizedStart,
          lineStarts,
        );
      } else {
        sourceSegments.push(segment);
      }
    }
  };
  for (const node of nodes) {
    switch (node.type) {
      case "text": {
        const value = node.value ?? "";
        const nodeStart = node.position?.start.offset;
        const nodeEnd = node.position?.end.offset;
        const exact =
          nodeStart !== undefined &&
          nodeEnd !== undefined &&
          markdown.slice(nodeStart, nodeEnd) === value;
        if (
          value === "" ||
          nodeStart === undefined ||
          nodeEnd === undefined ||
          !exact
        ) {
          append(
            value,
            node,
            { bold, italic, strikethrough },
            exact ? nodeStart : undefined,
          );
          break;
        }
        // remark only emits footnoteReference nodes when a matching
        // definition exists; scan the escape-preserving source slice so
        // escaped `\[^x]` stays literal while real references are collected.
        const source = markdown.slice(nodeStart, nodeEnd);
        const footnotePattern = /(?<!\\)\[\^([^\]]+)\]/g;
        let match: RegExpExecArray | null;
        let valueCursor = 0;
        while ((match = footnotePattern.exec(source))) {
          const remaining = value.slice(valueCursor);
          const at = remaining.indexOf(match[0]);
          if (at < 0) break;
          append(
            remaining.slice(0, at),
            node,
            { bold, italic, strikethrough },
            nodeStart + valueCursor,
          );
          const label = normalizeFootnoteLabel(match[0].slice(2, -1));
          refs.push(label);
          append(
            "⁎",
            node,
            { bold, italic, strikethrough, footnoteId: label },
            nodeStart + valueCursor + at,
          );
          valueCursor += at + match[0].length;
        }
        append(
          value.slice(valueCursor),
          node,
          { bold, italic, strikethrough },
          nodeStart + valueCursor,
        );
        break;
      }
      case "strong":
        appendNested(
          inline(
            node.children ?? [],
            markdown,
            lineStarts,
            true,
            italic,
            refs,
            strikethrough,
          ),
        );
        break;
      case "emphasis":
        appendNested(
          inline(
            node.children ?? [],
            markdown,
            lineStarts,
            bold,
            true,
            refs,
            strikethrough,
          ),
        );
        break;
      case "link": {
        const url = node.url ?? "";
        if (!permittedLink(url))
          throw new AgentDocxError(
            "REFERENCE_INVALID",
            `Unsupported link target: ${url}`,
            { position: pos(node) as unknown as Record<string, never> },
          );
        const nested = inline(
          node.children ?? [],
          markdown,
          lineStarts,
          bold,
          italic,
          refs,
          strikethrough,
        );
        for (const run of nested.runs)
          run.link = {
            target: url,
            ...(node.title ? { title: node.title } : {}),
          };
        appendNested(nested);
        break;
      }
      case "delete":
        appendNested(
          inline(
            node.children ?? [],
            markdown,
            lineStarts,
            bold,
            italic,
            refs,
            true,
          ),
        );
        break;
      case "break":
        append("\n", node, { bold, italic, strikethrough });
        break;
      case "footnoteReference": {
        const identifier = node.identifier;
        if (!identifier) return unsupported(node);
        const id = normalizeFootnoteLabel(identifier);
        refs.push(id);
        append("⁎", node, { bold, italic, strikethrough, footnoteId: id });
        break;
      }
      case "inlineCode":
        append(node.value ?? "", node, {
          bold,
          italic,
          strikethrough,
          literal: true,
        });
        break;
      case "image":
      case "html":
        unsupported(node);
        break;
      default:
        unsupported(node);
    }
  }
  return { runs, normalizedText, sourceSegments, footnoteRefs: [...refs] };
}
export function normalizeMarkdown(markdown: string): NormalizedDocument {
  assertXml10Legal(markdown);
  const root = unified()
    .use(remarkParse)
    .use(remarkGfm)
    .use(remarkFrontmatter, ["yaml"])
    .use(remarkMath, { singleDollarTextMath: false })
    .parse(markdown) as unknown as Node;
  const lineStarts = lineStartsFor(markdown);
  const blocks: FlowBlock[] = [];
  const footnotes = new Map<string, FootnoteDefinition>();
  const paragraphs: TextFlowBlock[] = [];
  const textBlock = (
    node: Node,
    kind: TextBlockKind,
    level?: number,
  ): TextFlowBlock => {
    const refs: string[] = [];
    const normalized = inline(
      node.children ?? [],
      markdown,
      lineStarts,
      false,
      false,
      refs,
    );
    return {
      kind,
      ...normalized,
      position: pos(node),
      ...(level === undefined ? {} : { level }),
      footnoteRefs: refs,
    };
  };
  const emit = (node: Node, kind: TextBlockKind, level?: number) => {
    const block = textBlock(node, kind, level);
    blocks.push(block);
    if (kind === "paragraph") paragraphs.push(block);
  };
  const firstDescendant = (
    node: Node,
    predicate: (candidate: Node) => boolean,
  ): Node | undefined => {
    const pending = [node];
    while (pending.length > 0) {
      const candidate = pending.pop()!;
      if (predicate(candidate)) return candidate;
      for (
        let index = (candidate.children?.length ?? 0) - 1;
        index >= 0;
        index--
      )
        pending.push(candidate.children![index]!);
    }
    return undefined;
  };
  const visit = (node: Node, context?: "blockquote" | "list", depth = 0) => {
    if (depth > MAX_NESTING_DEPTH)
      throw new AgentDocxError(
        "UNSUPPORTED_MARKDOWN",
        `Markdown nesting exceeds ${MAX_NESTING_DEPTH} levels`,
        { position: pos(node) as unknown as Record<string, never> },
      );
    switch (node.type) {
      case "paragraph":
        emit(node, context ?? "paragraph");
        break;
      case "heading":
        emit(node, "heading", node.depth);
        break;
      case "blockquote":
        for (const child of node.children ?? [])
          visit(child, "blockquote", depth + 1);
        break;
      case "list":
        for (const item of node.children ?? []) visit(item, "list", depth + 1);
        break;
      case "listItem":
        for (const child of node.children ?? [])
          visit(child, context ?? "list", depth + 1);
        break;
      case "footnoteDefinition": {
        const id = node.identifier
          ? normalizeFootnoteLabel(node.identifier)
          : undefined;
        if (!id || footnotes.has(id))
          throw new AgentDocxError(
            "UNSUPPORTED_MARKDOWN",
            `Missing or duplicate footnote: ${id ?? ""}`,
            { position: pos(node) as unknown as Record<string, never> },
          );
        if ((node.children?.length ?? 0) === 0) unsupported(node);
        const definitionBlocks: TextFlowBlock[] = [];
        for (const child of node.children ?? []) {
          if (child.type !== "paragraph") unsupported(child);
          definitionBlocks.push(textBlock(child, "footnote"));
        }
        footnotes.set(id, {
          id,
          position: pos(node),
          blocks: definitionBlocks,
          footnoteRefs: definitionBlocks.flatMap((block) => block.footnoteRefs),
        });
        break;
      }
      case "table": {
        const sourceRows = node.children ?? [];
        const width = sourceRows.reduce(
          (maximum, row) => Math.max(maximum, row.children?.length ?? 0),
          0,
        );
        const alignments = Array.from(
          { length: width },
          (_, columnIndex) => node.align?.[columnIndex] ?? null,
        );
        const rows = sourceRows.map((row) => {
          if (row.type !== "tableRow") return unsupported(row);
          return Array.from({ length: width }, (_, columnIndex): TableCell => {
            const cell = row.children?.[columnIndex];
            if (!cell)
              return {
                runs: [],
                normalizedText: "",
                sourceSegments: [],
                footnoteRefs: [],
                position: pos(row),
                alignment: alignments[columnIndex] ?? null,
              };
            const offendingReference = firstDescendant(
              cell,
              (candidate) =>
                candidate.type === "image" || candidate.type === "html",
            );
            if (offendingReference) unsupported(offendingReference);
            const refs: string[] = [];
            const normalized = inline(
              cell.children ?? [],
              markdown,
              lineStarts,
              false,
              false,
              refs,
            );
            return {
              ...normalized,
              position: pos(cell),
              alignment: alignments[columnIndex] ?? null,
            };
          });
        });
        blocks.push({
          kind: "table",
          position: pos(node),
          alignments,
          rows,
        });
        break;
      }
      case "thematicBreak":
        blocks.push({ kind: "thematic-break", position: pos(node) });
        break;
      case "html":
        if ((node.value ?? "").trim() === "<!-- pagebreak -->")
          blocks.push({ kind: "pagebreak", position: pos(node) });
        else unsupported(node);
        break;
      case "code":
      case "yaml":
      case "image":
      case "math":
        unsupported(node);
        break;
      default:
        unsupported(node);
    }
  };
  for (const child of root.children ?? []) visit(child);
  const visited = new Set<string>();
  const validateDefinition = (id: string, source: TextFlowBlock) => {
    const pending: Array<{ id: string; source: TextFlowBlock }> = [
      { id, source },
    ];
    while (pending.length > 0) {
      const current = pending.pop()!;
      const definition = footnotes.get(current.id);
      if (!definition)
        throw new AgentDocxError(
          "UNSUPPORTED_MARKDOWN",
          `Missing footnote definition: ${current.id}`,
          {
            position: current.source.position as unknown as Record<
              string,
              never
            >,
          },
        );
      if (visited.has(current.id)) continue;
      visited.add(current.id);
      for (const block of definition.blocks)
        for (const nested of block.footnoteRefs)
          pending.push({ id: nested, source: block });
    }
  };
  for (const block of blocks) {
    if (
      block.kind !== "table" &&
      block.kind !== "thematic-break" &&
      block.kind !== "pagebreak"
    )
      for (const id of block.footnoteRefs) validateDefinition(id, block);
    if (block.kind === "table")
      for (const row of block.rows)
        for (const cell of row)
          for (const id of cell.footnoteRefs)
            validateDefinition(id, {
              kind: "paragraph",
              runs: cell.runs,
              normalizedText: cell.normalizedText,
              sourceSegments: cell.sourceSegments,
              position: cell.position,
              footnoteRefs: cell.footnoteRefs,
            });
  }
  for (const [id, definition] of footnotes)
    for (const block of definition.blocks) validateDefinition(id, block);
  return { blocks, footnotes, paragraphs };
}
