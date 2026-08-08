import type { SaxesTagNS } from "saxes";
import { isBlockId, type BlockId } from "../legal/model.js";
import type { DocxFidelityItem } from "./contracts.js";
import { docxXmlAttribute, parseDocxXml } from "./package.js";
import { isWordprocessingElement } from "./opc.js";
import { unsupported } from "./helpers.js";
import { AgentDocxError } from "../types.js";

export type Paragraph = {
  bookmark: BlockId | null;
  bookmarkName: string | null;
  text: string;
  heading: number | null;
  sourcePart: string;
  comments: readonly TrackedCommentAnchor[];
};

export type NativeRevision = {
  id: string;
  kind: "ins" | "del" | "moveFrom" | "moveTo";
  author: string;
  date: string | null;
  text: string;
};

export type TrackedCommentAnchor = {
  id: string;
  start: number;
  end: number;
};

export type TrackedParagraph = {
  bookmark: BlockId | null;
  baseText: string;
  headText: string;
  visibleText: string;
  heading: number | null;
  sourcePart: string;
  revisions: readonly NativeRevision[];
  comments: readonly TrackedCommentAnchor[];
};

export type TrackedMaterial = {
  baseSource: string;
  headSource: string;
  paragraphs: readonly TrackedParagraph[];
};
export const parseTrackedParagraphs = (
  xml: string,
  sourcePart: string,
): {
  paragraphs: readonly TrackedParagraph[];
  unsupported: readonly DocxFidelityItem<"unsupported">[];
} => {
  const paragraphs: TrackedParagraph[] = [];
  const unsupportedItems: DocxFidelityItem<"unsupported">[] = [];
  let current: {
    bookmark: BlockId | null;
    baseText: string;
    headText: string;
    visibleText: string;
    heading: number | null;
    revisions: NativeRevision[];
    comments: TrackedCommentAnchor[];
  } | null = null;
  let revision: NativeRevision | null = null;
  const commentStarts: Array<{ id: string; start: number }> = [];
  let inText = false;
  const append = (value: string): void => {
    if (!current) return;
    if (revision) {
      revision.text += value;
      if (revision.kind === "ins" || revision.kind === "moveTo") {
        current.headText += value;
        current.visibleText += value;
      } else current.baseText += value;
      return;
    }
    current.baseText += value;
    current.headText += value;
    current.visibleText += value;
  };
  parseDocxXml(
    xml,
    (tag) => {
      if (!isWordprocessingElement(tag)) return;
      if (
        reportUnsupportedConstruct(tag, "tracked", sourcePart, unsupportedItems)
      )
        return;
      if (tag.local === "p") {
        if (current) unsupported("Nested DOCX paragraphs are unsupported");
        current = {
          bookmark: null,
          baseText: "",
          headText: "",
          visibleText: "",
          heading: null,
          revisions: [],
          comments: [],
        };
      }
      if (!current) {
        if (["ins", "del", "moveFrom", "moveTo"].includes(tag.local))
          unsupported("Tracked revision is outside a paragraph");
        return;
      }
      if (tag.local === "bookmarkStart") {
        const bookmark = fromBookmarkName(docxXmlAttribute(tag, "name"));
        if (bookmark) {
          if (current.bookmark)
            unsupported("Paragraph has multiple block bookmarks");
          current.bookmark = bookmark;
        }
        return;
      }
      if (tag.local === "pStyle") {
        const style = docxXmlAttribute(tag, "val") ?? "";
        const match = /^Heading([1-6])$/.exec(style);
        if (match) current.heading = Number(match[1]);
        return;
      }
      if (tag.local === "commentRangeStart") {
        const id = docxXmlAttribute(tag, "id");
        if (!id || !/^\d+$/.test(id) || commentStarts.length > 0)
          unsupported("Comment range nesting is malformed");
        commentStarts.push({
          id: id as string,
          start: current.visibleText.length,
        });
        return;
      }
      if (["ins", "del", "moveFrom", "moveTo"].includes(tag.local)) {
        if (revision) unsupported("Nested tracked revisions are unsupported");
        const id = docxXmlAttribute(tag, "id");
        const author = docxXmlAttribute(tag, "author");
        const date = docxXmlAttribute(tag, "date");
        if (!id || !/^\d+$/.test(id) || author === undefined)
          unsupported("Tracked revision has invalid native attribution");
        if (date !== undefined && Number.isNaN(new Date(date).valueOf()))
          unsupported("Tracked revision has an invalid native date");
        revision = {
          id: id as string,
          kind: tag.local as NativeRevision["kind"],
          author: author as string,
          date: date ?? null,
          text: "",
        };
        return;
      }
      if (tag.local === "t" || tag.local === "delText") inText = true;
      else if (tag.local === "tab") append("\t");
      else if (tag.local === "br" || tag.local === "cr") append("\n");
    },
    (tag) => {
      if (!isWordprocessingElement(tag)) return;
      if (tag.local === "t" || tag.local === "delText") {
        inText = false;
        return;
      }
      if (tag.local === "commentRangeEnd") {
        const id = docxXmlAttribute(tag, "id");
        const start = commentStarts.pop();
        if (!id || !start || start.id !== id)
          unsupported("Comment range nesting is malformed");
        const safeStart = start as { id: string; start: number };
        current!.comments.push({
          id: safeStart.id,
          start: safeStart.start,
          end: current!.visibleText.length,
        });
        return;
      }
      if (["ins", "del", "moveFrom", "moveTo"].includes(tag.local)) {
        const closedRevision = revision;
        if (!closedRevision || closedRevision.kind !== tag.local)
          unsupported("Tracked revision nesting is malformed");
        current!.revisions.push(closedRevision as NativeRevision);
        revision = null;
        return;
      }
      if (tag.local === "p") {
        if (revision || commentStarts.length > 0)
          unsupported(
            "Tracked revision or comment crosses a paragraph boundary",
          );
        paragraphs.push({ ...current!, sourcePart });
        current = null;
      }
    },
    (text) => {
      if (inText) append(text);
    },
  );
  if (current || revision || inText)
    unsupported("Tracked DOCX text nesting is malformed");
  return { paragraphs, unsupported: unsupportedItems };
};
export const escapedMarkdown = (text: string): string =>
  text
    .replaceAll("\\", "\\\\")
    .replaceAll("`", "\\`")
    .replaceAll("*", "\\*")
    .replaceAll("_", "\\_")
    .replaceAll("[", "\\[")
    .replaceAll("]", "\\]")
    .replaceAll("<", "\\<")
    .replaceAll(">", "\\>");

export const fromBookmarkName = (name: string | undefined): BlockId | null => {
  if (!name?.startsWith("adx_")) return null;
  const hex = name.slice("adx_".length);
  if (!/^[0-9a-f]{32}$/.test(hex)) return null;
  const candidate = `b_${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20)}`;
  return isBlockId(candidate) ? candidate : null;
};

const isGeneratedBodyBookmark = (name: string): boolean =>
  /^adx_body_\d{6}$/.test(name);

export const emittedBookmarkName = (name: string | undefined): string | null =>
  name !== undefined &&
  (fromBookmarkName(name) !== null || isGeneratedBodyBookmark(name))
    ? name
    : null;

const UNSUPPORTED_OXML_CONSTRUCTS: Record<
  "clean" | "tracked",
  ReadonlySet<string>
> = {
  clean: new Set(["tbl", "ins", "del", "moveFrom", "moveTo", "altChunk"]),
  tracked: new Set(["tbl", "altChunk"]),
};

const reportUnsupportedConstruct = (
  tag: SaxesTagNS,
  mode: "clean" | "tracked",
  sourcePart: string,
  items: DocxFidelityItem<"unsupported">[],
): boolean => {
  if (!UNSUPPORTED_OXML_CONSTRUCTS[mode].has(tag.local)) return false;
  items.push({
    status: "unsupported",
    partPath: sourcePart,
    relationshipId: null,
    ooxmlKind: `w:${tag.local}`,
    count: 1,
    blockIds: [],
    sourcePositions: [],
    explanation:
      "This OOXML construct cannot be represented by the version 1 Markdown importer.",
  });
  return true;
};

export const parseParagraphs = (
  xml: string,
  sourcePart: string,
): {
  paragraphs: readonly Paragraph[];
  unsupported: readonly DocxFidelityItem<"unsupported">[];
} => {
  const paragraphs: Paragraph[] = [];
  const unsupportedItems: DocxFidelityItem<"unsupported">[] = [];
  let current: {
    bookmark: BlockId | null;
    bookmarkName: string | null;
    text: string;
    heading: number | null;
    comments: TrackedCommentAnchor[];
  } | null = null;
  const commentStarts: Array<{ id: string; start: number }> = [];
  let runText = "";
  let inText = false;
  let unsupportedDepth = 0;
  parseDocxXml(
    xml,
    (tag: SaxesTagNS) => {
      if (!isWordprocessingElement(tag)) return;
      if (
        reportUnsupportedConstruct(tag, "clean", sourcePart, unsupportedItems)
      )
        unsupportedDepth++;
      if (tag.local === "p")
        current = {
          bookmark: null,
          bookmarkName: null,
          text: "",
          heading: null,
          comments: [],
        };
      if (tag.local === "commentRangeStart" && current) {
        const id = docxXmlAttribute(tag, "id");
        if (!id || !/^\d+$/.test(id) || commentStarts.length > 0)
          unsupported("Comment range nesting is malformed");
        commentStarts.push({ id: id as string, start: current.text.length });
      }
      if (tag.local === "bookmarkStart" && current) {
        const name = emittedBookmarkName(docxXmlAttribute(tag, "name"));
        if (name !== null) {
          if (current.bookmarkName !== null)
            unsupported("Paragraph has multiple agent-docx bookmarks");
          current.bookmarkName = name;
          current.bookmark = fromBookmarkName(name);
        }
      }
      if (tag.local === "pStyle" && current) {
        const style = docxXmlAttribute(tag, "val") ?? "";
        const match = /^Heading([1-6])$/.exec(style);
        if (match) current.heading = Number(match[1]);
      }
      if (tag.local === "r") runText = "";
      if (tag.local === "t" && current) inText = true;
      if (tag.local === "tab" && current) current.text += "\t";
      if ((tag.local === "br" || tag.local === "cr") && current)
        current.text += "\n";
    },
    (tag) => {
      if (!isWordprocessingElement(tag)) return;
      if (tag.local === "t") inText = false;
      if (tag.local === "r" && current) current.text += runText;
      if (tag.local === "commentRangeEnd" && current) {
        const id = docxXmlAttribute(tag, "id");
        const start = commentStarts.pop();
        if (!id || !start || id !== start.id)
          unsupported("Comment range nesting is malformed");
        const safeStart = start as { id: string; start: number };
        current.comments.push({
          id: safeStart.id,
          start: safeStart.start,
          end: current.text.length,
        });
      }
      if (tag.local === "p" && current) {
        if (commentStarts.length > 0)
          unsupported("Comment range crosses a paragraph boundary");
        paragraphs.push({ ...current, sourcePart });
        current = null;
      }
      if (
        reportUnsupportedConstruct(tag, "clean", sourcePart, unsupportedItems)
      )
        unsupportedDepth--;
    },
    (text) => {
      if (inText && current) runText += text;
    },
  );
  if (unsupportedDepth !== 0 || commentStarts.length !== 0)
    throw new AgentDocxError(
      "DOCX_IMPORT_UNSUPPORTED",
      "Malformed OOXML nesting",
    );
  return { paragraphs, unsupported: unsupportedItems };
};
