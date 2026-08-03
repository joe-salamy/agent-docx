import {
  Bookmark,
  CommentRangeEnd,
  CommentRangeStart,
  CommentReference,
  DeletedTextRun,
  Document,
  InsertedTextRun,
  Packer,
  Paragraph,
  TextRun,
  type ParagraphChild,
} from "docx";
import type { LayoutProfile, TextStyle } from "../types.js";
import type {
  DocumentChrome,
  LegalBlock,
  LegalDocument,
  LitigationMetadata,
  ReviewAnnotation,
} from "../legal/model.js";
import type { Change, ChangeSet } from "../revisions/types.js";
import { blockBookmark } from "../legal/model.js";
import {
  addSemanticManifest,
  createNativeDocumentChrome,
  nativeSectionProperties,
  nativeStyles,
  normalizeGeneratedPackage,
  numbering,
  paragraphOptions,
} from "./generate.js";

export type GeneratedRedlineDocx = {
  bytes: Uint8Array;
  revisionCount: number;
  commentCount: number;
  bodyParagraphs: readonly {
    id: string;
    index: number;
    position: LegalBlock["position"];
    preview: string;
  }[];
};

export type GenerateRedlineDocxOptions = {
  chrome?: DocumentChrome;
  metadata?: LitigationMetadata;
  pageCount?: number;
  semanticManifest?: Readonly<Record<string, unknown>>;
  createdAt?: string;
};

type TextBlock = Extract<
  LegalBlock,
  | { kind: "paragraph" }
  | { kind: "heading" }
  | { kind: "blockquote" }
  | { kind: "numbered-paragraph" }
>;

const textBlock = (block: LegalBlock): block is TextBlock =>
  block.kind === "paragraph" ||
  block.kind === "heading" ||
  block.kind === "blockquote" ||
  block.kind === "numbered-paragraph";

const blockText = (block: TextBlock): string =>
  block.runs.map((run) => `${run.text}${run.hardBreakAfter ? "\n" : ""}`).join("");

const textForLegalBlock = (block: LegalBlock): string => {
  if (textBlock(block)) return blockText(block);
  if (block.kind === "list")
    return block.items
      .map((item) =>
        item.paragraphs
          .map((paragraph) =>
            paragraph.runs
              .map((run) => `${run.text}${run.hardBreakAfter ? "\n" : ""}`)
              .join(""),
          )
          .join("\n"),
      )
      .join("\n");
  if (block.kind === "table")
    return block.rows
      .map((row) =>
        row
          .map((cell) =>
            cell.paragraphs
              .map((paragraph) =>
                paragraph.runs
                  .map((run) => `${run.text}${run.hardBreakAfter ? "\n" : ""}`)
                  .join(""),
              )
              .join("\n"),
          )
          .join(" | "),
      )
      .join("\n");
  if (block.kind === "exhibit")
    return [block.label, ...block.blocks.map(textForLegalBlock)]
      .filter(Boolean)
      .join("\n");
  if (block.kind === "length-exclusion")
    return block.blocks.map(textForLegalBlock).join("\n");
  if (block.kind === "image") return `[Image: ${block.alt}]`;
  return block.sourceText;
};

const textBlockForRedline = (block: LegalBlock): TextBlock => {
  if (textBlock(block)) return block;
  return {
    id: block.id,
    kind: "paragraph",
    position: block.position,
    sourceText: block.sourceText,
    segments: block.segments,
    runs: [{
      text: textForLegalBlock(block),
      bold: false,
      italic: false,
      strikethrough: false,
      literal: false,
      hardBreakAfter: false,
    }],
    footnoteRefs: [],
  };
};

const styleFor = (block: TextBlock, profile: LayoutProfile): TextStyle =>
  block.kind === "heading"
    ? profile.headings[String(block.level) as "1"]
    : block.kind === "blockquote"
      ? profile.blockquote
      : block.kind === "numbered-paragraph"
        ? profile.list
        : profile.body;

const paragraphStyleId = (block: TextBlock): string =>
  block.kind === "heading"
    ? `AgentDocxHeading${block.level}`
    : block.kind === "blockquote"
      ? "AgentDocxBlockQuote"
      : block.kind === "numbered-paragraph"
        ? "AgentDocxList"
        : "AgentDocxBody";

const revision = (change: Change, id: number) => {
  const createdAt = change.attribution.createdAt;
  const date =
    createdAt === null
      ? undefined
      : (() => {
          const parsed = new Date(createdAt);
          if (Number.isNaN(parsed.valueOf()))
            throw new Error(`Invalid revision attribution date: ${createdAt}`);
          return parsed.toISOString();
        })();
  return {
    id,
    author: change.attribution.author?.name ?? "",
    ...(date === undefined ? {} : { date }),
  } as { id: number; author: string; date: string };
};

const textRun = (text: string, style: TextStyle, profile: LayoutProfile): TextRun =>
  new TextRun({
    text,
    font: profile.requestedFontFamily,
    size: style.fontSizeTwips / 10,
    bold: style.bold,
    italics: style.italic,
  });

const insertedRun = (
  text: string,
  style: TextStyle,
  profile: LayoutProfile,
  change: Change,
  id: number,
): InsertedTextRun =>
  new InsertedTextRun({
    ...revision(change, id),
    text,
    font: profile.requestedFontFamily,
    size: style.fontSizeTwips / 10,
    bold: style.bold,
    italics: style.italic,
  });

const deletedRun = (
  text: string,
  style: TextStyle,
  profile: LayoutProfile,
  change: Change,
  id: number,
): DeletedTextRun =>
  new DeletedTextRun({
    ...revision(change, id),
    text,
    font: profile.requestedFontFamily,
    size: style.fontSizeTwips / 10,
    bold: style.bold,
    italics: style.italic,
  });

const changeByBlock = (changeSet: ChangeSet): ReadonlyMap<string, Change> =>
  new Map(
    changeSet.changes
      .filter((change): change is Extract<Change, { blockId: string }> => "blockId" in change)
      .map((change) => [change.blockId, change]),
  );

const deletedBefore = (
  base: readonly TextBlock[],
  headIds: ReadonlySet<string>,
  changes: ReadonlyMap<string, Change>,
): ReadonlyMap<string | null, TextBlock[]> => {
  const result = new Map<string | null, TextBlock[]>();
  for (let index = 0; index < base.length; index++) {
    const block = base[index]!;
    if (headIds.has(block.id) || changes.get(block.id)?.kind !== "delete-block") continue;
    const next = base.slice(index + 1).find((candidate) => headIds.has(candidate.id));
    const key = next?.id ?? null;
    const entries = result.get(key) ?? [];
    entries.push(block);
    result.set(key, entries);
  }
  return result;
};

type RedlineComment = {
  id: number;
  annotation: ReviewAnnotation;
};

const commentInitials = (author: string): string =>
  author
    .split(/\s+/)
    .filter(Boolean)
    .map((part) => part[0]!.toUpperCase())
    .join("")
    .slice(0, 8);

const isCodePointBoundary = (text: string, offset: number): boolean =>
  offset >= 0 &&
  offset <= text.length &&
  (offset === 0 ||
    offset === text.length ||
    !(
      text.charCodeAt(offset - 1) >= 0xd800 &&
      text.charCodeAt(offset - 1) <= 0xdbff &&
      text.charCodeAt(offset) >= 0xdc00 &&
      text.charCodeAt(offset) <= 0xdfff
    ));

const commentsFor = (
  annotations: readonly ReviewAnnotation[],
  blocks: readonly TextBlock[],
): readonly RedlineComment[] => {
  const blockIds = new Set(blocks.map((block) => block.id));
  return annotations
    .filter(
      (annotation) =>
        annotation.status === "open" && blockIds.has(annotation.blockId),
    )
    .sort((left, right) => left.id.localeCompare(right.id))
    .map((annotation, id) => ({ id, annotation }));
};

const commentChildren = (
  text: string,
  style: TextStyle,
  profile: LayoutProfile,
  comments: readonly RedlineComment[],
): readonly ParagraphChild[] => {
  const ranged = comments.filter(
    (comment): comment is RedlineComment & {
      annotation: ReviewAnnotation & { range: { start: number; end: number } };
    } => comment.annotation.range !== undefined,
  );
  if (ranged.length === 0) return [textRun(text, style, profile)];
  const ordered = [...ranged].sort(
    (left, right) =>
      left.annotation.range.start - right.annotation.range.start ||
      left.annotation.range.end - right.annotation.range.end ||
      left.id - right.id,
  );
  let cursor = 0;
  const children: ParagraphChild[] = [];
  for (const comment of ordered) {
    const { start, end } = comment.annotation.range;
    if (
      start < cursor ||
      end > text.length ||
      start >= end ||
      !isCodePointBoundary(text, start) ||
      !isCodePointBoundary(text, end)
    )
      throw new Error(
        "Open review comment ranges overlap, split a code point, or exceed their block text",
      );
    if (start > cursor)
      children.push(textRun(text.slice(cursor, start), style, profile));
    children.push(new CommentRangeStart(comment.id));
    children.push(textRun(text.slice(start, end), style, profile));
    children.push(new CommentRangeEnd(comment.id));
    cursor = end;
  }
  if (cursor < text.length)
    children.push(textRun(text.slice(cursor), style, profile));
  return children;
};

const redlineBlocks = (document: LegalDocument): readonly TextBlock[] => [
  ...document.blocks.map(textBlockForRedline),
  ...document.footnotes.map((footnote): TextBlock => ({
    id: footnote.id,
    kind: "paragraph",
    position: footnote.position,
    sourceText: footnote.sourceText,
    segments: footnote.segments,
    runs: [{
      text: footnote.paragraphs
        .map((paragraph) =>
          paragraph.runs
            .map((run) => `${run.text}${run.hardBreakAfter ? "\n" : ""}`)
            .join(""),
        )
        .join("\n"),
      bold: false,
      italic: false,
      strikethrough: false,
      literal: false,
      hardBreakAfter: false,
    }],
    footnoteRefs: [],
  })),
];

/**
 * Emits OOXML `w:ins` and `w:del` runs for source-mapped leaf-block changes.
 * The visible document is always the selected head; deleted content remains in
 * the package for Word's review view.
 */
export const generateRedlineDocx = async (
  base: LegalDocument,
  head: LegalDocument,
  changeSet: ChangeSet,
  profile: LayoutProfile,
  options: GenerateRedlineDocxOptions = {},
): Promise<GeneratedRedlineDocx> => {
  const baseBlocks = redlineBlocks(base);
  const headBlocks = redlineBlocks(head);
  const changes = changeByBlock(changeSet);
  const deleted = deletedBefore(
    baseBlocks,
    new Set(headBlocks.map((block) => block.id)),
    changes,
  );
  const comments = commentsFor(head.annotations, headBlocks);
  const commentsByBlock = new Map<string, RedlineComment[]>();
  for (const comment of comments) {
    const entries = commentsByBlock.get(comment.annotation.blockId) ?? [];
    entries.push(comment);
    commentsByBlock.set(comment.annotation.blockId, entries);
  }
  let revisionId = 1;
  const bodyParagraphs: Array<GeneratedRedlineDocx["bodyParagraphs"][number]> = [];
  const children: Paragraph[] = [];
  const appendDeleted = (block: TextBlock) => {
    const change = changes.get(block.id);
    if (!change || change.kind !== "delete-block") return;
    const style = styleFor(block, profile);
    children.push(
      new Paragraph({
        style: paragraphStyleId(block),
        ...paragraphOptions(style, profile.pagination.widowOrphanControl),
        children: [
          new Bookmark({
            id: blockBookmark(block.id),
            children: [
              deletedRun(blockText(block), style, profile, change, revisionId++),
            ],
          }),
        ],
      }),
    );
  };
  for (const block of headBlocks) {
    for (const removed of deleted.get(block.id) ?? []) appendDeleted(removed);
    const style = styleFor(block, profile);
    const change = changes.get(block.id);
    const currentText = blockText(block);
    let childrenForBlock: readonly ParagraphChild[];
    if (change?.kind === "insert-block")
      childrenForBlock = [insertedRun(currentText, style, profile, change, revisionId++)];
    else if (
      change?.kind === "replace-text" ||
      change?.kind === "replace-block" ||
      change?.kind === "move-block"
    ) {
      const previous =
        change.kind === "replace-text"
          ? change.oldText
          : change.kind === "replace-block"
            ? change.oldBlock.kind === "footnote"
              ? change.oldBlock.paragraphs
                  .map((paragraph) =>
                    paragraph.runs.map((run) => run.text).join(""),
                  )
                  .join("\n")
              : textForLegalBlock(change.oldBlock)
            : currentText;
      childrenForBlock = [
        deletedRun(previous, style, profile, change, revisionId++),
        insertedRun(currentText, style, profile, change, revisionId++),
      ];
    } else childrenForBlock = [textRun(currentText, style, profile)];
    bodyParagraphs.push({
      id: blockBookmark(block.id),
      index: bodyParagraphs.length,
      position: block.position,
      preview: currentText.replace(/\s+/g, " ").trim().slice(0, 80),
    });
    const blockComments = commentsByBlock.get(block.id) ?? [];
    const rangedComments = blockComments.filter(
      (comment) => comment.annotation.range !== undefined,
    );
    if (rangedComments.length > 0) {
      if (change)
        throw new Error(
          "Native redline comments with text ranges cannot target revised blocks",
        );
      childrenForBlock = commentChildren(
        currentText,
        style,
        profile,
        rangedComments,
      );
    }
    const blockWideComments = blockComments.filter(
      (comment) => comment.annotation.range === undefined,
    );
    const bookmark = new Bookmark({
      id: blockBookmark(block.id),
      children: childrenForBlock,
    });
    children.push(
      new Paragraph({
        style: paragraphStyleId(block),
        ...paragraphOptions(style, profile.pagination.widowOrphanControl),
        children: [
          ...blockWideComments.map(
            (comment) => new CommentRangeStart(comment.id),
          ),
          bookmark,
          ...[...blockWideComments]
            .reverse()
            .map((comment) => new CommentRangeEnd(comment.id)),
          ...blockComments.map((comment) => new CommentReference(comment.id)),
        ],
      }),
    );
  }
  for (const removed of deleted.get(null) ?? []) appendDeleted(removed);
  const commentDefinitions = comments.map(({ id, annotation }) => {
    const date =
      annotation.createdAt === null
        ? undefined
        : (() => {
            const parsed = new Date(annotation.createdAt);
            if (Number.isNaN(parsed.valueOf()))
              throw new Error(`Invalid comment date: ${annotation.createdAt}`);
            return parsed;
          })();
    const author = annotation.author?.name ?? "";
    return {
      id,
      author,
      initials: commentInitials(author),
      ...(date === undefined ? {} : { date }),
      resolved: false,
      children: [new Paragraph({ children: [new TextRun(annotation.message)] })],
    };
  });
  const chrome = options.chrome ?? head.chrome;
  const nativeChrome = createNativeDocumentChrome(
    profile,
    chrome,
    options.metadata ?? head.metadata,
    options.pageCount,
  );
  const document = new Document({
    styles: { paragraphStyles: nativeStyles(profile) },
    numbering: numbering(profile),
    features: { trackRevisions: true, updateFields: true },
    evenAndOddHeaderAndFooters:
      nativeChrome.evenAndOddHeaderAndFooters,
    ...(commentDefinitions.length
      ? { comments: { children: commentDefinitions } }
      : {}),
    sections: [{
      properties: nativeSectionProperties(profile, chrome, nativeChrome),
      ...(Object.keys(nativeChrome.headers).length > 0
        ? { headers: nativeChrome.headers }
        : {}),
      ...(Object.keys(nativeChrome.footers).length > 0
        ? { footers: nativeChrome.footers }
        : {}),
      children: children.length === 0 ? [new Paragraph({})] : children,
    }],
  });
  const packed = await Packer.toBuffer(document);
  const bytes = options.semanticManifest
    ? await addSemanticManifest(
        packed,
        {
          ...options.semanticManifest,
          emittedBlocks: bodyParagraphs.map(({ id, index }) => ({
            bookmark: id,
            index,
          })),
        },
        options.createdAt,
      )
    : await normalizeGeneratedPackage(packed, options.createdAt);
  return {
    bytes,
    revisionCount: revisionId - 1,
    commentCount: comments.length,
    bodyParagraphs,
  };
};
