import { unified } from "unified";
import remarkParse from "remark-parse";
import remarkGfm from "remark-gfm";
import {
  AgentDocxError,
  type SectionHeading,
  type SourcePosition,
} from "./types.js";
type Node = {
  type: string;
  value?: string;
  depth?: number;
  ordered?: boolean;
  children?: Node[];
  position?: {
    start: { line: number; column: number; offset?: number };
    end: { line: number; column: number; offset?: number };
  };
  identifier?: string;
  label?: string;
  url?: string;
};
export type InlineRun = {
  text: string;
  bold: boolean;
  italic: boolean;
  footnoteId?: string;
};
export type FlowBlock = {
  kind:
    | "paragraph"
    | "heading"
    | "blockquote"
    | "list"
    | "footnote"
    | "pagebreak";
  runs: InlineRun[];
  position: SourcePosition;
  level?: number;
  hardBreakAfter: boolean;
  footnoteRefs: string[];
};
export type NormalizedDocument = {
  blocks: FlowBlock[];
  footnotes: Map<string, FlowBlock>;
  paragraphs: FlowBlock[];
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
const unsupported = (node: Node): never => {
  throw new AgentDocxError(
    "UNSUPPORTED_MARKDOWN",
    `Unsupported Markdown node: ${node.type}`,
    { position: pos(node) as unknown as Record<string, never> },
  );
};
function inline(
  nodes: Node[],
  bold = false,
  italic = false,
  refs: string[] = [],
): InlineRun[] {
  const out: InlineRun[] = [];
  for (const node of nodes) {
    switch (node.type) {
      case "text": {
        const value = node.value ?? "";
        let cursor = 0;
        for (const match of value.matchAll(/\[\^([^\]]+)\]/g)) {
          const start = match.index ?? 0;
          if (start > cursor)
            out.push({ text: value.slice(cursor, start), bold, italic });
          const id = match[1]!.trim().toLowerCase();
          refs.push(id);
          out.push({ text: "⁎", bold, italic, footnoteId: id });
          cursor = start + match[0].length;
        }
        if (cursor < value.length)
          out.push({ text: value.slice(cursor), bold, italic });
        break;
      }
      case "strong":
        out.push(...inline(node.children ?? [], true, italic, refs));
        break;
      case "emphasis":
        out.push(...inline(node.children ?? [], bold, true, refs));
        break;
      case "link":
      case "delete":
        out.push(...inline(node.children ?? [], bold, italic, refs));
        break;
      case "break":
        out.push({ text: "\n", bold, italic });
        break;
      case "footnoteReference":
        if (!node.identifier) unsupported(node);
        refs.push(node.identifier!);
        out.push({ text: "⁎", bold, italic, footnoteId: node.identifier! });
        break;
      case "inlineCode":
      case "image":
      case "html":
        unsupported(node);
        break;
      default:
        unsupported(node);
    }
  }
  return out;
}
const emptyPos: SourcePosition = {
  start: { line: 1, column: 1, offset: 0 },
  end: { line: 1, column: 1, offset: 0 },
};
export function normalizeMarkdown(markdown: string): NormalizedDocument {
  const root = unified()
    .use(remarkParse)
    .use(remarkGfm)
    .parse(markdown) as unknown as Node;
  const blocks: FlowBlock[] = [];
  const footnotes = new Map<string, FlowBlock>();
  const paragraphs: FlowBlock[] = [];
  const emit = (node: Node, kind: FlowBlock["kind"], level?: number) => {
    const refs: string[] = [];
    const b: FlowBlock = {
      kind,
      runs: inline(node.children ?? [], false, false, refs),
      position: pos(node),
      hardBreakAfter: false,
      footnoteRefs: refs,
    };
    if (level !== undefined) b.level = level;
    blocks.push(b);
    if (kind === "paragraph") paragraphs.push(b);
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
        for (const c of node.children ?? []) visit(c, "blockquote");
        break;
      case "list":
        for (const item of node.children ?? []) visit(item, "list");
        break;
      case "listItem":
        for (const c of node.children ?? []) visit(c, context ?? "list");
        break;
      case "footnoteDefinition": {
        const id = node.identifier;
        if (!id || footnotes.has(id))
          throw new AgentDocxError(
            "UNSUPPORTED_MARKDOWN",
            `Missing or duplicate footnote: ${id ?? ""}`,
          );
        if (
          (node.children?.length ?? 0) !== 1 ||
          node.children?.[0]?.type !== "paragraph"
        )
          unsupported(node);
        const child = node.children![0]!;
        const refs: string[] = [];
        footnotes.set(id, {
          kind: "footnote",
          runs: inline(child.children ?? [], false, false, refs),
          position: pos(node),
          hardBreakAfter: false,
          footnoteRefs: refs,
        });
        break;
      }
      case "html":
        if ((node.value ?? "").trim() === "<!-- pagebreak -->")
          blocks.push({
            kind: "pagebreak",
            runs: [],
            position: pos(node),
            hardBreakAfter: true,
            footnoteRefs: [],
          });
        else unsupported(node);
        break;
      case "thematicBreak":
        unsupported(node);
        break;
      case "table":
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
  const validateDefinition = (id: string, source: FlowBlock) => {
    const definition = footnotes.get(id);
    if (!definition)
      throw new AgentDocxError(
        "UNSUPPORTED_MARKDOWN",
        `Missing footnote definition: ${id}`,
        { position: source.position as unknown as Record<string, never> },
      );
    if (visited.has(id)) return;
    visited.add(id);
    for (const nested of definition.footnoteRefs)
      validateDefinition(nested, definition);
  };
  for (const block of blocks)
    for (const id of block.footnoteRefs) validateDefinition(id, block);
  for (const [id, definition] of footnotes) validateDefinition(id, definition);
  return { blocks, footnotes, paragraphs };
}
export const emptyPosition = emptyPos;
