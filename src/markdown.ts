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
import type {
  AuthorityReference,
  LegalBlock,
} from "./legal/model.js";
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
};

function inline(
  nodes: Node[],
  markdown: string,
  lineStarts: readonly number[],
  bold = false,
  italic = false,
  refs: string[] = [],
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
        const exact =
          nodeStart !== undefined &&
          node.position?.end.offset !== undefined &&
          markdown.slice(nodeStart, node.position.end.offset) === value;
        let cursor = 0;
        for (const match of value.matchAll(/\[\^([^\]]+)\]/g)) {
          const start = match.index ?? 0;
          if (start > cursor)
            append(
              value.slice(cursor, start),
              node,
              { bold, italic },
              exact ? nodeStart + cursor : undefined,
            );
          const id = match[1]!.trim().toLowerCase();
          refs.push(id);
          append("⁎", node, { bold, italic, footnoteId: id });
          cursor = start + match[0].length;
        }
        if (cursor < value.length)
          append(
            value.slice(cursor),
            node,
            { bold, italic },
            exact ? nodeStart + cursor : undefined,
          );
        break;
      }
      case "strong":
        appendNested(
          inline(node.children ?? [], markdown, lineStarts, true, italic, refs),
        );
        break;
      case "emphasis":
        appendNested(
          inline(node.children ?? [], markdown, lineStarts, bold, true, refs),
        );
        break;
      case "link":
      case "delete":
        appendNested(
          inline(node.children ?? [], markdown, lineStarts, bold, italic, refs),
        );
        break;
      case "break":
        append("\n", node, { bold, italic });
        break;
      case "footnoteReference": {
        if (!node.identifier) unsupported(node);
        const id = node.identifier!.toLowerCase();
        refs.push(id);
        append("⁎", node, { bold, italic, footnoteId: id });
        break;
      }
      case "inlineCode":
        append(node.value ?? "", node, { bold, italic, literal: true });
        break;
      case "image":
      case "html":
        unsupported(node);
        break;
      default:
        unsupported(node);
    }
  }
  return { runs, normalizedText, sourceSegments };
}
export function normalizeMarkdown(markdown: string): NormalizedDocument {
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
    if (predicate(node)) return node;
    for (const child of node.children ?? []) {
      const found = firstDescendant(child, predicate);
      if (found) return found;
    }
    return undefined;
  };
  const visit = (node: Node, context?: "blockquote" | "list") => {
    switch (node.type) {
      case "paragraph":
        emit(node, context ?? "paragraph");
        break;
      case "heading":
        emit(node, "heading", node.depth);
        break;
      case "blockquote":
        for (const child of node.children ?? []) visit(child, "blockquote");
        break;
      case "list":
        for (const item of node.children ?? []) visit(item, "list");
        break;
      case "listItem":
        for (const child of node.children ?? [])
          visit(child, context ?? "list");
        break;
      case "footnoteDefinition": {
        const id = node.identifier?.toLowerCase();
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
        const alignments = node.align ?? [];
        const rows = (node.children ?? []).map((row) =>
          (row.children ?? []).map((cell, columnIndex): TableCell => {
            const offendingReference = firstDescendant(
              cell,
              (candidate) =>
                candidate.type === "footnoteReference" ||
                candidate.type === "image" ||
                candidate.type === "html",
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
            if (refs.length) unsupported(cell);
            return {
              ...normalized,
              position: pos(cell),
              alignment: alignments[columnIndex] ?? null,
            };
          }),
        );
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
    const definition = footnotes.get(id);
    if (!definition)
      throw new AgentDocxError(
        "UNSUPPORTED_MARKDOWN",
        `Missing footnote definition: ${id}`,
        { position: source.position as unknown as Record<string, never> },
      );
    if (visited.has(id)) return;
    visited.add(id);
    for (const block of definition.blocks)
      for (const nested of block.footnoteRefs)
        validateDefinition(nested, block);
  };
  for (const block of blocks)
    if (
      block.kind !== "table" &&
      block.kind !== "thematic-break" &&
      block.kind !== "pagebreak"
    )
      for (const id of block.footnoteRefs) validateDefinition(id, block);
  for (const [id, definition] of footnotes)
    for (const block of definition.blocks) validateDefinition(id, block);
  return { blocks, footnotes, paragraphs };
}
