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
import type { LayoutProfile, TextStyle } from "../layout/profile.js";
import { AgentDocxError } from "../types.js";
import type { TextFlowBlock } from "../markdown.js";
import { flowStyleFor } from "../layout/style.js";
import type {
  DocumentChrome,
  LegalBlock,
  LegalDocument,
  LitigationMetadata,
  ReviewAnnotation,
} from "../legal/model.js";
import type { Change, ChangeSet } from "../revisions/types.js";
import { visibleRangeForSource } from "../revisions/diff.js";
import { visibleTextForBlock } from "../legal/visible-text.js";
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

const REDLINE_SUPPORTED_KINDS = new Set<string>([
  "paragraph",
  "heading",
  "blockquote",
  "numbered-paragraph",
  "list",
  "footnote",
  "exhibit",
  "length-exclusion",
]);

const assertRedlineSupported = (document: LegalDocument): void => {
  const visit = (blocks: readonly LegalBlock[]): void => {
    for (const block of blocks) {
      if (!REDLINE_SUPPORTED_KINDS.has(block.kind))
        throw new AgentDocxError(
          "DOCX_REDLINE_UNSUPPORTED",
          `Redline does not support ${block.kind} blocks (block ${block.id})`,
        );
      if (block.kind === "exhibit" || block.kind === "length-exclusion")
        visit(block.blocks);
    }
  };
  visit(document.blocks);
};

const flowView = (block: TextBlock): TextFlowBlock => ({
  kind: block.kind === "numbered-paragraph" ? "list" : block.kind,
  runs: [...block.runs],
  normalizedText: "",
  sourceSegments: [],
  position: block.position,
  footnoteRefs: [],
  ...(block.kind === "heading" ? { level: block.level } : {}),
  ...(block.kind === "numbered-paragraph"
    ? { legalKind: "numbered-paragraph" as const, numberedLevel: block.level - 1 }
    : {}),
});

const styleFor = (block: TextBlock, profile: LayoutProfile): TextStyle =>
  flowStyleFor(flowView(block), profile).style;

const paragraphStyleId = (block: TextBlock, profile: LayoutProfile): string =>
  flowStyleFor(flowView(block), profile).styleId;

const revision = (change: Change, id: number) => {
  const createdAt = change.attribution.createdAt;
  const date =
    createdAt === null
      ? undefined
      : (() => {
          const parsed = new Date(createdAt);
          if (Number.isNaN(parsed.valueOf()))
            throw new AgentDocxError(
              "DOCX_REDLINE_UNSUPPORTED",
              `Invalid revision attribution date: ${createdAt}`,
            );
          return parsed.toISOString();
        })();
  return {
    id,
    author: change.attribution.author?.name ?? "",
    ...(date === undefined ? {} : { date }),
  } as { id: number; author: string; date: string };
};

const textRun = (
  text: string,
  style: TextStyle,
  profile: LayoutProfile,
): TextRun =>
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

type BlockChangeMap = ReadonlyMap<string, readonly Change[]>;

const changesByBlock = (changeSet: ChangeSet): BlockChangeMap => {
  const result = new Map<string, Change[]>();
  for (const change of changeSet.changes) {
    if (!("blockId" in change)) continue;
    const changes = result.get(change.blockId) ?? [];
    changes.push(change);
    result.set(change.blockId, changes);
  }
  return result;
};

const structuralChange = (
  changes: readonly Change[],
): Extract<
  Change,
  {
    kind:
      | "insert-block"
      | "delete-block"
      | "move-block"
      | "replace-block"
      | "replace-container-shell";
  }
> | null =>
  changes.find(
    (
      change,
    ): change is Extract<
      Change,
      {
        kind:
          | "insert-block"
          | "delete-block"
          | "move-block"
          | "replace-block"
          | "replace-container-shell";
      }
    > =>
      change.kind === "insert-block" ||
      change.kind === "delete-block" ||
      change.kind === "move-block" ||
      change.kind === "replace-block" ||
      change.kind === "replace-container-shell",
  ) ?? null;

const deletedBefore = (
  base: readonly TextBlock[],
  headIds: ReadonlySet<string>,
  changes: BlockChangeMap,
): ReadonlyMap<string | null, TextBlock[]> => {
  const result = new Map<string | null, TextBlock[]>();
  for (let index = 0; index < base.length; index++) {
    const block = base[index]!;
    if (
      headIds.has(block.id) ||
      !changes.get(block.id)?.some((change) => change.kind === "delete-block")
    )
      continue;
    const next = base
      .slice(index + 1)
      .find((candidate) => headIds.has(candidate.id));
    const key = next?.id ?? null;
    const entries = result.get(key) ?? [];
    entries.push(block);
    result.set(key, entries);
  }
  return result;
};
type TextChange = Extract<
  Change,
  { kind: "insert-text" | "delete-text" | "replace-text" }
>;

type VisibleTextEdit = {
  change: TextChange;
  oldStart: number;
  oldEnd: number;
  newStart: number;
  newEnd: number;
};

const redlineTextChildren = (
  base: TextBlock,
  head: TextBlock,
  changes: readonly TextChange[],
  style: TextStyle,
  profile: LayoutProfile,
  nextRevision: () => number,
): readonly ParagraphChild[] | null => {
  const oldText = visibleTextForBlock(base);
  const newText = visibleTextForBlock(head);
  const edits: VisibleTextEdit[] = [];
  for (const change of changes) {
    const oldRange =
      change.kind === "insert-text"
        ? visibleRangeForSource(base, change.oldOffset, change.oldOffset)
        : visibleRangeForSource(
            base,
            change.oldSource.start,
            change.oldSource.end,
          );
    const newRange =
      change.kind === "delete-text"
        ? visibleRangeForSource(head, change.newOffset, change.newOffset)
        : visibleRangeForSource(
            head,
            change.newSource.start,
            change.newSource.end,
          );
    if (!oldRange || !newRange) return null;
    if (
      oldRange.start < 0 ||
      oldRange.end > oldText.length ||
      newRange.start < 0 ||
      newRange.end > newText.length
    )
      return null;
    const oldSlice = oldText.slice(oldRange.start, oldRange.end);
    const newSlice = newText.slice(newRange.start, newRange.end);
    if (
      (change.kind === "insert-text" && oldSlice !== "") ||
      (change.kind === "delete-text" && newSlice !== "") ||
      (change.kind !== "insert-text" &&
        change.kind !== "delete-text" &&
        oldSlice !== change.oldText) ||
      (change.kind !== "delete-text" &&
        change.kind !== "insert-text" &&
        newSlice !== change.newText) ||
      (change.kind === "insert-text" && newSlice !== change.newText) ||
      (change.kind === "delete-text" && oldSlice !== change.oldText)
    )
      return null;
    edits.push({
      change,
      oldStart: oldRange.start,
      oldEnd: oldRange.end,
      newStart: newRange.start,
      newEnd: newRange.end,
    });
  }
  edits.sort(
    (left, right) =>
      left.oldStart - right.oldStart ||
      left.newStart - right.newStart ||
      left.change.id.localeCompare(right.change.id),
  );
  const children: ParagraphChild[] = [];
  let oldCursor = 0;
  let newCursor = 0;
  const appendText = (text: string): void => {
    if (text.length > 0) children.push(textRun(text, style, profile));
  };
  for (const edit of edits) {
    if (
      edit.oldStart < oldCursor ||
      edit.newStart < newCursor ||
      edit.oldEnd < edit.oldStart ||
      edit.newEnd < edit.newStart
    )
      return null;
    appendText(newText.slice(newCursor, edit.newStart));
    if (edit.oldEnd > edit.oldStart)
      children.push(
        deletedRun(
          oldText.slice(edit.oldStart, edit.oldEnd),
          style,
          profile,
          edit.change,
          nextRevision(),
        ),
      );
    if (edit.newEnd > edit.newStart)
      children.push(
        insertedRun(
          newText.slice(edit.newStart, edit.newEnd),
          style,
          profile,
          edit.change,
          nextRevision(),
        ),
      );
    oldCursor = edit.oldEnd;
    newCursor = edit.newEnd;
  }
  appendText(newText.slice(newCursor));
  if (oldCursor > oldText.length || newCursor > newText.length) return null;
  return children;
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

import { isCodePointBoundary } from "./text.js";
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
    (
      comment,
    ): comment is RedlineComment & {
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
      throw new AgentDocxError(
        "DOCX_INVALID",
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

const redlineBlocks = (document: LegalDocument): readonly TextBlock[] => {
  const asText = (block: LegalBlock): TextBlock =>
    textBlock(block)
      ? block
      : {
          id: block.id,
          kind: "paragraph",
          position: block.position,
          sourceText: block.sourceText,
          segments: block.segments,
          runs: [
            {
              text: visibleTextForBlock(block),
              bold: false,
              italic: false,
              strikethrough: false,
              literal: false,
              hardBreakAfter: false,
            },
          ],
          footnoteRefs: [],
        };
  return [
    ...document.blocks.map(asText),
    ...document.footnotes.map(
      (footnote): TextBlock => ({
        id: footnote.id,
        kind: "paragraph",
        position: footnote.position,
        sourceText: footnote.sourceText,
        segments: footnote.segments,
        runs: [
          {
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
          },
        ],
        footnoteRefs: [],
      }),
    ),
  ];
};

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
  assertRedlineSupported(base);
  assertRedlineSupported(head);
  const baseBlocks = redlineBlocks(base);
  const headBlocks = redlineBlocks(head);
  const baseById = new Map(baseBlocks.map((block) => [block.id, block]));
  const changes = changesByBlock(changeSet);
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
  const bodyParagraphs: Array<GeneratedRedlineDocx["bodyParagraphs"][number]> =
    [];
  const children: Paragraph[] = [];
  const appendDeleted = (block: TextBlock) => {
    const change = changes
      .get(block.id)
      ?.find(
        (entry): entry is Extract<Change, { kind: "delete-block" }> =>
          entry.kind === "delete-block",
      );
    if (!change) return;
    const style = styleFor(block, profile);
    children.push(
      new Paragraph({
        style: paragraphStyleId(block, profile),
        ...paragraphOptions(style, profile.pagination.widowOrphanControl),
        children: [
          new Bookmark({
            id: blockBookmark(block.id),
            children: [
              deletedRun(
                visibleTextForBlock(block),
                style,
                profile,
                change,
                revisionId++,
              ),
            ],
          }),
        ],
      }),
    );
  };
  for (const block of headBlocks) {
    for (const removed of deleted.get(block.id) ?? []) appendDeleted(removed);
    const style = styleFor(block, profile);
    const blockChanges = changes.get(block.id) ?? [];
    const change = structuralChange(blockChanges);
    const textChanges = blockChanges.filter(
      (entry): entry is TextChange =>
        entry.kind === "insert-text" ||
        entry.kind === "delete-text" ||
        entry.kind === "replace-text",
    );
    const currentText = visibleTextForBlock(block);
    let childrenForBlock: readonly ParagraphChild[];
    if (change?.kind === "insert-block")
      childrenForBlock = [
        insertedRun(currentText, style, profile, change, revisionId++),
      ];
    else if (
      change?.kind === "replace-block" ||
      change?.kind === "move-block"
    ) {
      const previous =
        change.kind === "replace-block"
          ? change.oldBlock.kind === "footnote"
            ? change.oldBlock.paragraphs
                .map((paragraph) =>
                  paragraph.runs.map((run) => run.text).join(""),
                )
                .join("\n")
            : visibleTextForBlock(change.oldBlock)
          : currentText;
      childrenForBlock = [
        deletedRun(previous, style, profile, change, revisionId++),
        insertedRun(currentText, style, profile, change, revisionId++),
      ];
    } else if (change?.kind === "replace-container-shell") {
      const previous = baseById.get(block.id);
      if (!previous)
        throw new AgentDocxError(
          "DOCX_INVALID",
          `Missing base block for container shell replacement: ${block.id}`,
        );
      childrenForBlock = [
        deletedRun(visibleTextForBlock(previous), style, profile, change, revisionId++),
        insertedRun(currentText, style, profile, change, revisionId++),
      ];
    } else if (textChanges.length > 0 && baseById.get(block.id)) {
      childrenForBlock = redlineTextChildren(
        baseById.get(block.id)!,
        block,
        textChanges,
        style,
        profile,
        () => revisionId++,
      ) ?? [
        deletedRun(
          visibleTextForBlock(baseById.get(block.id)!),
          style,
          profile,
          textChanges[0]!,
          revisionId++,
        ),
        insertedRun(currentText, style, profile, textChanges[0]!, revisionId++),
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
      if (blockChanges.length > 0)
        throw new AgentDocxError(
          "DOCX_INVALID",
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
        style: paragraphStyleId(block, profile),
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
              throw new AgentDocxError(
                "DOCX_REDLINE_UNSUPPORTED",
                `Invalid comment date: ${annotation.createdAt}`,
              );
            return parsed;
          })();
    const author = annotation.author?.name ?? "";
    return {
      id,
      author,
      initials: commentInitials(author),
      ...(date === undefined ? {} : { date }),
      resolved: false,
      children: [
        new Paragraph({ children: [new TextRun(annotation.message)] }),
      ],
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
    evenAndOddHeaderAndFooters: nativeChrome.evenAndOddHeaderAndFooters,
    ...(commentDefinitions.length
      ? { comments: { children: commentDefinitions } }
      : {}),
    sections: [
      {
        properties: nativeSectionProperties(profile, chrome, nativeChrome),
        ...(Object.keys(nativeChrome.headers).length > 0
          ? { headers: nativeChrome.headers }
          : {}),
        ...(Object.keys(nativeChrome.footers).length > 0
          ? { footers: nativeChrome.footers }
          : {}),
        children: children.length === 0 ? [new Paragraph({})] : children,
      },
    ],
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
