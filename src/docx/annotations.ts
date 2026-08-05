import { createHash } from "node:crypto";
import { AgentDocxError } from "../types.js";
import { parseLegalMarkdown } from "../legal/parse.js";
import {
  blockBookmark,
  type BlockId,
  type DocumentChrome,
  type LegalDocument,
  type LitigationMetadata,
  type ReviewAnnotation,
} from "../legal/model.js";
import { visibleTextForBlock } from "../legal/visible-text.js";
import { isCodePointBoundary } from "./text.js";
import type {
  SemanticManifest,
  SemanticRevisionMapEntry,
} from "./manifest.js";
import type {
  Paragraph,
  TrackedMaterial,
  TrackedParagraph,
} from "./tracked.js";
import type { ImportedAsset } from "./attachments.js";
import { docxXmlAttribute, parseDocxXml } from "./package.js";
const unsupported = (message: string): never => {
  throw new AgentDocxError("DOCX_IMPORT_UNSUPPORTED", message);
};
const asObject = (value: unknown, label: string): Record<string, unknown> => {
  if (value === null || typeof value !== "object" || Array.isArray(value))
    unsupported(`${label} must be an object`);
  return value as Record<string, unknown>;
};

export type NativeComment = {
  id: string;
  author: string;
  date: string | null;
  text: string;
};
export const reconstructTrackedMaterial = (
  semantic: SemanticManifest,
  paragraphs: readonly TrackedParagraph[],
  assets: Readonly<Record<string, ImportedAsset>>,
): TrackedMaterial => {
  if (semantic.mode !== "redline")
    unsupported("Tracked reconstruction requires a redline semantic manifest");
  const projectedDocument = asObject(
    semantic.document,
    "Semantic manifest document",
  );
  const parsed = parseLegalMarkdown(semantic.source, {
    projectId: semantic.projectId,
    documentId: semantic.documentId,
    metadata: projectedDocument.metadata as LitigationMetadata,
    chrome: (projectedDocument.chrome ?? {}) as DocumentChrome,
    assets,
    requireMarkers: true,
  }).document;
  const byBookmark = new Map<
    string,
    LegalDocument["blocks"][number] | LegalDocument["footnotes"][number]
  >();
  for (const block of [...parsed.blocks, ...parsed.footnotes])
    byBookmark.set(blockBookmark(block.id), block);
  const replacements: Array<{ start: number; end: number; source: string }> =
    [];
  for (const paragraph of paragraphs) {
    if (paragraph.revisions.length === 0) continue;
    const bookmark = paragraph.bookmark;
    if (!bookmark)
      unsupported("Tracked revision has no agent-docx block bookmark");
    const block = byBookmark.get(blockBookmark(bookmark as BlockId));
    if (!block)
      unsupported(
        "Tracked revision bookmark is not declared by the semantic source",
      );
    const resolvedBlock = block as
      | LegalDocument["blocks"][number]
      | LegalDocument["footnotes"][number];
    if (visibleTextForBlock(resolvedBlock) !== paragraph.headText)
      unsupported("Tracked revision head text does not match semantic source");
    if (paragraph.baseText === paragraph.headText)
      unsupported("Tracked revision has no visible base-to-head change");
    const occurrences =
      resolvedBlock.sourceText.split(paragraph.headText).length - 1;
    if (occurrences !== 1)
      unsupported(
        "Tracked revision cannot be mapped unambiguously to a Markdown source range",
      );
    replacements.push({
      start: resolvedBlock.position.start.offset,
      end: resolvedBlock.position.end.offset,
      source: resolvedBlock.sourceText.replace(
        paragraph.headText,
        paragraph.baseText,
      ),
    });
  }
  if (replacements.length === 0) {
    if (semantic.revisionMap.length > 0)
      unsupported(
        "Redline semantic manifest declares unrepresented tracked revisions",
      );
    return {
      baseSource: semantic.source,
      headSource: semantic.source,
      paragraphs,
    };
  }
  replacements.sort((left, right) => right.start - left.start);
  for (let index = 1; index < replacements.length; index++)
    if (replacements[index - 1]!.start < replacements[index]!.end)
      unsupported("Tracked revisions overlap in the semantic source");
  let baseSource = semantic.source;
  for (const replacement of replacements)
    baseSource =
      baseSource.slice(0, replacement.start) +
      replacement.source +
      baseSource.slice(replacement.end);
  return {
    baseSource,
    headSource: semantic.source,
    paragraphs,
  };
};
export const trackedAnnotations = (
  semantic: SemanticManifest,
  tracked: TrackedMaterial,
  document: LegalDocument,
  nativeComments: ReadonlyMap<string, NativeComment>,
): readonly ReviewAnnotation[] => {
  const anchors = tracked.paragraphs.flatMap((paragraph) =>
    paragraph.comments.map((anchor) => ({ paragraph, anchor })),
  );
  if (anchors.length !== nativeComments.size)
    unsupported("Native comment definitions and anchors do not match");
  const anchorsById = new Map<string, (typeof anchors)[number]>();
  for (const entry of anchors) {
    if (anchorsById.has(entry.anchor.id))
      unsupported("Native comment has multiple or duplicate anchors");
    anchorsById.set(entry.anchor.id, entry);
  }
  const nativeIds = [...nativeComments.keys()].sort(
    (left, right) => Number(left) - Number(right),
  );
  if (semantic.commentMap.length !== nativeIds.length)
    unsupported("Semantic comment map does not match native comment count");
  const blocks = new Map(
    [...document.blocks, ...document.footnotes].map((block) => [
      block.id,
      block,
    ]),
  );
  return nativeIds.map((nativeId, index) => {
    const native = nativeComments.get(nativeId)!;
    const mapped = semantic.commentMap[index]!;
    const candidate = anchorsById.get(nativeId);
    if (!candidate || !candidate.paragraph.bookmark)
      unsupported("Native comment does not anchor to an agent-docx block");
    const anchor = candidate as (typeof anchors)[number];
    const bookmark = anchor.paragraph.bookmark as BlockId;
    const block = blocks.get(bookmark);
    if (!block) unsupported("Native comment anchors an unknown block");
    const resolvedBlock = block as
      | LegalDocument["blocks"][number]
      | LegalDocument["footnotes"][number];
    const visible = visibleTextForBlock(resolvedBlock);
    if (
      anchor.anchor.end > visible.length ||
      anchor.anchor.start > anchor.anchor.end ||
      !isCodePointBoundary(visible, anchor.anchor.start) ||
      !isCodePointBoundary(visible, anchor.anchor.end)
    )
      unsupported("Native comment range is not code-point safe");
    const blockWide =
      anchor.anchor.start === 0 && anchor.anchor.end === visible.length;
    if (mapped.blockWide !== blockWide)
      unsupported("Semantic comment map does not match native comment anchor");
    if (!/^a_[0-9a-f-]{36}$/.test(mapped.annotationId))
      unsupported("Semantic comment map has an invalid annotation ID");
    return {
      id: mapped.annotationId as `a_${string}`,
      blockId: resolvedBlock.id,
      ...(blockWide
        ? {}
        : { range: { start: anchor.anchor.start, end: anchor.anchor.end } }),
      author:
        native.author === ""
          ? null
          : {
              name: native.author,
              ...(mapped.authorEmail === null
                ? {}
                : { email: mapped.authorEmail }),
            },
      createdAt: native.date,
      message: native.text,
      status: "open" as const,
    };
  });
};
export const redlineAnnotations = (
  semantic: SemanticManifest,
  tracked: TrackedMaterial,
  document: LegalDocument,
  nativeComments: ReadonlyMap<string, NativeComment>,
  visibleTextByBlock: ReadonlyMap<BlockId, string>,
): readonly ReviewAnnotation[] => {
  const anchors = tracked.paragraphs.flatMap((paragraph) =>
    paragraph.comments.map((anchor) => ({ paragraph, anchor })),
  );
  if (anchors.length !== nativeComments.size)
    unsupported("Native comment definitions and anchors do not match");
  const anchorsById = new Map<string, (typeof anchors)[number]>();
  for (const entry of anchors) {
    if (anchorsById.has(entry.anchor.id))
      unsupported("Native comment has multiple or duplicate anchors");
    anchorsById.set(entry.anchor.id, entry);
  }
  const nativeIds = [...nativeComments.keys()].sort(
    (left, right) => Number(left) - Number(right),
  );
  if (nativeIds.length < semantic.commentMap.length)
    unsupported("Semantic comment map has missing native comments");
  const declaredIds = nativeIds.slice(0, semantic.commentMap.length);
  const highestDeclaredId =
    declaredIds.length === 0
      ? -1
      : Math.max(...declaredIds.map((id) => Number(id)));
  if (
    nativeIds
      .slice(semantic.commentMap.length)
      .some((id) => Number(id) <= highestDeclaredId)
  )
    unsupported(
      "Additional reviewer comments cannot be distinguished from declared comments",
    );
  const blocks = new Map(
    [...document.blocks, ...document.footnotes].map((block) => [
      block.id,
      block,
    ]),
  );
  const usedAnnotationIds = new Set(
    semantic.commentMap.map((entry) => entry.annotationId),
  );
  return nativeIds.map((nativeId, index) => {
    const native = nativeComments.get(nativeId)!;
    const candidate = anchorsById.get(nativeId);
    if (!candidate || !candidate.paragraph.bookmark)
      unsupported("Native comment does not anchor to an agent-docx block");
    const anchor = candidate as (typeof anchors)[number];
    const bookmark = anchor.paragraph.bookmark as BlockId;
    const block = blocks.get(bookmark);
    if (!block) unsupported("Native comment anchors an unknown block");
    const resolvedBlock = block as
      | LegalDocument["blocks"][number]
      | LegalDocument["footnotes"][number];
    const visible =
      visibleTextByBlock.get(bookmark) ?? visibleTextForBlock(resolvedBlock);
    if (
      anchor.anchor.end > visible.length ||
      anchor.anchor.start > anchor.anchor.end ||
      !isCodePointBoundary(visible, anchor.anchor.start) ||
      !isCodePointBoundary(visible, anchor.anchor.end)
    )
      unsupported("Native comment range is not code-point safe");
    const blockWide =
      anchor.anchor.start === 0 && anchor.anchor.end === visible.length;
    const mapped = semantic.commentMap[index];
    if (mapped) {
      if (mapped.blockWide !== blockWide)
        unsupported(
          "Semantic comment map does not match native comment anchor",
        );
      if (!/^a_[0-9a-f-]{36}$/.test(mapped.annotationId))
        unsupported("Semantic comment map has an invalid annotation ID");
      return {
        id: mapped.annotationId as `a_${string}`,
        blockId: resolvedBlock.id,
        ...(blockWide
          ? {}
          : { range: { start: anchor.anchor.start, end: anchor.anchor.end } }),
        author:
          native.author === ""
            ? null
            : {
                name: native.author,
                ...(mapped.authorEmail === null
                  ? {}
                  : { email: mapped.authorEmail }),
              },
        createdAt: native.date,
        message: native.text,
        status: "open" as const,
      };
    }
    const digest = createHash("sha256")
      .update(
        `agent-docx:review-comment:${semantic.projectId}:${semantic.documentId}:${bookmark}:${nativeId}`,
      )
      .digest("hex");
    const annotationId = `a_${digest.slice(0, 8)}-${digest.slice(
      8,
      12,
    )}-${digest.slice(12, 16)}-${digest.slice(16, 20)}-${digest.slice(
      20,
      32,
    )}` as `a_${string}`;
    if (usedAnnotationIds.has(annotationId))
      unsupported("Additional reviewer comment has a duplicate annotation ID");
    usedAnnotationIds.add(annotationId);
    return {
      id: annotationId,
      blockId: resolvedBlock.id,
      ...(blockWide
        ? {}
        : { range: { start: anchor.anchor.start, end: anchor.anchor.end } }),
      author: native.author === "" ? null : { name: native.author },
      createdAt: native.date,
      message: native.text,
      status: "open" as const,
    };
  });
};

export type RedlineDecision = "accept" | "reject";
export type RedlineDecisionMap = Readonly<Record<string, RedlineDecision>>;

export const extendedRevisionEntry = (
  entry: SemanticRevisionMapEntry,
): entry is SemanticRevisionMapEntry & {
  blockId: BlockId | null;
  baseText: string;
  headText: string;
} =>
  entry.blockId !== undefined &&
  entry.baseText !== undefined &&
  entry.headText !== undefined;

export const semanticBlocks = (
  document: LegalDocument,
): ReadonlyMap<
  BlockId,
  LegalDocument["blocks"][number] | LegalDocument["footnotes"][number]
> => {
  const blocks = new Map<
    BlockId,
    LegalDocument["blocks"][number] | LegalDocument["footnotes"][number]
  >();
  const visit = (entries: readonly LegalDocument["blocks"][number][]): void => {
    for (const block of entries) {
      blocks.set(block.id, block);
      if (block.kind === "exhibit" || block.kind === "length-exclusion")
        visit(block.blocks);
    }
  };
  visit(document.blocks);
  for (const footnote of document.footnotes) blocks.set(footnote.id, footnote);
  return blocks;
};

const replaceVisibleBlockText = (
  block: LegalDocument["blocks"][number] | LegalDocument["footnotes"][number],
  visibleText: string,
): { start: number; end: number; replacement: string } | null => {
  const current = visibleTextForBlock(block);
  if (current === visibleText) return null;
  const occurrences = block.sourceText.split(current).length - 1;
  if (current.length === 0 || occurrences !== 1)
    unsupported(
      "Resolved redline paragraph cannot be mapped unambiguously to its Markdown source range",
    );
  return {
    start: block.position.start.offset,
    end: block.position.end.offset,
    replacement: block.sourceText.replace(current, visibleText),
  };
};

export const sourceWithVisibleBlocks = (
  source: string,
  blocks: ReadonlyMap<
    BlockId,
    LegalDocument["blocks"][number] | LegalDocument["footnotes"][number]
  >,
  visibleByBlock: ReadonlyMap<BlockId, string>,
): string => {
  const replacements = [...visibleByBlock.entries()].flatMap(
    ([blockId, text]) => {
      const block = blocks.get(blockId);
      if (!block)
        unsupported("Resolved redline paragraph references an unknown block");
      const replacement = replaceVisibleBlockText(
        block as
          | LegalDocument["blocks"][number]
          | LegalDocument["footnotes"][number],
        text,
      );
      return replacement ? [replacement] : [];
    },
  );
  replacements.sort((left, right) => right.start - left.start);
  for (let index = 1; index < replacements.length; index++)
    if (replacements[index - 1]!.start < replacements[index]!.end)
      unsupported("Resolved redline paragraph source ranges overlap");
  let result = source;
  for (const replacement of replacements)
    result =
      result.slice(0, replacement.start) +
      replacement.replacement +
      result.slice(replacement.end);
  return result;
};

const applyRevisionGroup = (
  headText: string,
  entries: readonly (SemanticRevisionMapEntry & {
    blockId: BlockId;
    baseText: string;
    headText: string;
  })[],
  mask: number,
): string | null => {
  const replacements: Array<{
    start: number;
    end: number;
    replacement: string;
  }> = [];
  for (const [index, entry] of entries.entries()) {
    if ((mask & (1 << index)) === 0) continue;
    if (entry.headText.length === 0) return null;
    const positions: number[] = [];
    let cursor = 0;
    while (true) {
      const found = headText.indexOf(entry.headText, cursor);
      if (found < 0) break;
      positions.push(found);
      cursor = found + entry.headText.length;
    }
    if (positions.length !== 1) return null;
    const start = positions[0]!;
    replacements.push({
      start,
      end: start + entry.headText.length,
      replacement: entry.baseText,
    });
  }
  replacements.sort((left, right) => right.start - left.start);
  for (let index = 1; index < replacements.length; index++)
    if (replacements[index - 1]!.start < replacements[index]!.end) return null;
  let result = headText;
  for (const replacement of replacements)
    result =
      result.slice(0, replacement.start) +
      replacement.replacement +
      result.slice(replacement.end);
  return result;
};

export const resolveRevisionGroups = (
  semantic: SemanticManifest,
  document: LegalDocument,
  tracked: TrackedMaterial,
): {
  decisions: RedlineDecisionMap;
  visibleByBlock: ReadonlyMap<BlockId, string>;
  baseByBlock: ReadonlyMap<BlockId, string>;
} => {
  const blocks = semanticBlocks(document);
  const byBookmark = new Map<BlockId, TrackedParagraph>();
  for (const paragraph of tracked.paragraphs) {
    if (!paragraph.bookmark) {
      if (paragraph.visibleText.length > 0)
        unsupported("Resolved redline text has no agent-docx block bookmark");
      continue;
    }
    if (byBookmark.has(paragraph.bookmark))
      unsupported("Resolved redline has duplicate agent-docx block bookmarks");
    byBookmark.set(paragraph.bookmark, paragraph);
  }
  const allExtended = semantic.revisionMap.every(extendedRevisionEntry);
  if (!allExtended) {
    for (const [blockId, paragraph] of byBookmark) {
      const block = semanticBlocks(document).get(blockId);
      if (!block)
        unsupported("Resolved redline paragraph references an unknown block");
      if (
        paragraph.visibleText !==
        visibleTextForBlock(
          block as
            | LegalDocument["blocks"][number]
            | LegalDocument["footnotes"][number],
        )
      )
        unsupported(
          "Resolved redline changes without base/head text are ambiguous",
        );
    }
    return {
      decisions: Object.fromEntries(
        semantic.revisionMap.map((entry) => [entry.changeId, "accept"]),
      ) as RedlineDecisionMap,
      visibleByBlock: new Map(
        [...byBookmark].map(([blockId, paragraph]) => [
          blockId,
          paragraph.visibleText,
        ]),
      ),
      baseByBlock: new Map(),
    };
  }
  const groups = new Map<
    BlockId,
    (SemanticRevisionMapEntry & {
      blockId: BlockId;
      baseText: string;
      headText: string;
    })[]
  >();
  const decisionsSet = new Map<string, RedlineDecision>();
  for (const entry of semantic.revisionMap) {
    if (entry.blockId === null)
      unsupported(
        "Resolved redline change cannot be attributed to an agent-docx block",
      );
    const blockId = entry.blockId as BlockId;
    if (!blocks.has(blockId) || !byBookmark.has(blockId))
      unsupported("Resolved redline change references an unknown block");
    const entries = groups.get(blockId) ?? [];
    entries.push({
      ...entry,
      blockId,
      baseText: entry.baseText as string,
      headText: entry.headText as string,
    });
    groups.set(blockId, entries);
  }
  const visibleByBlock = new Map<BlockId, string>();
  for (const [blockId, paragraph] of byBookmark)
    visibleByBlock.set(blockId, paragraph.visibleText);
  for (const [blockId, paragraph] of byBookmark) {
    const block = blocks.get(blockId);
    if (!block)
      unsupported("Resolved redline paragraph references an unknown block");
    const expectedHead = visibleTextForBlock(
      block as
        | LegalDocument["blocks"][number]
        | LegalDocument["footnotes"][number],
    );
    const entries = groups.get(blockId) ?? [];
    if (entries.length === 0 && paragraph.visibleText !== expectedHead)
      unsupported(
        "Resolved redline contains a foreign edit outside declared changes",
      );
    if (entries.length === 0) continue;
    if (paragraph.visibleText === expectedHead) {
      for (const entry of entries) decisionsSet.set(entry.changeId, "accept");
      continue;
    }
    if (!allExtended)
      unsupported(
        "Resolved redline changes without base/head text are ambiguous",
      );
    if (entries.length > 20)
      unsupported(
        "Resolved redline paragraph has too many changes to attribute safely",
      );
    const matches: Array<{ mask: number; text: string }> = [];
    const combinations = 1 << entries.length;
    for (let mask = 0; mask < combinations; mask++) {
      const candidate = applyRevisionGroup(expectedHead, entries, mask);
      if (candidate === paragraph.visibleText)
        matches.push({ mask, text: candidate });
    }
    if (matches.length === 0 && entries.length === 1) {
      const entry = entries[0]!;
      if (
        entry.headText.length === 0 &&
        paragraph.visibleText === entry.baseText
      )
        matches.push({ mask: 1, text: paragraph.visibleText });
    }
    if (matches.length !== 1)
      unsupported(
        matches.length === 0
          ? "Resolved redline paragraph does not match any declared accept/reject decision"
          : "Resolved redline paragraph matches multiple accept/reject decisions",
      );
    const mask = matches[0]!.mask;
    for (const [index, entry] of entries.entries())
      decisionsSet.set(
        entry.changeId,
        (mask & (1 << index)) === 0 ? "accept" : "reject",
      );
  }
  const baseByBlock = new Map<BlockId, string>();
  for (const [blockId, entries] of groups) {
    const block = blocks.get(blockId)!;
    const head = visibleTextForBlock(
      block as
        | LegalDocument["blocks"][number]
        | LegalDocument["footnotes"][number],
    );
    const candidate =
      allExtended && entries.length <= 20
        ? applyRevisionGroup(head, entries, (1 << entries.length) - 1)
        : null;
    if (candidate !== null) baseByBlock.set(blockId, candidate);
  }
  return {
    decisions: Object.fromEntries(decisionsSet) as RedlineDecisionMap,
    visibleByBlock,
    baseByBlock,
  };
};

export const cleanAnnotations = (
  paragraphs: readonly Paragraph[],
  document: LegalDocument,
  nativeComments: ReadonlyMap<string, NativeComment>,
): readonly ReviewAnnotation[] => {
  const anchors = paragraphs.flatMap((paragraph) =>
    paragraph.comments.map((anchor) => ({ paragraph, anchor })),
  );
  if (anchors.length !== nativeComments.size)
    unsupported("Native comment definitions and anchors do not match");
  const byNativeId = new Map<string, (typeof anchors)[number]>();
  for (const anchor of anchors) {
    if (byNativeId.has(anchor.anchor.id))
      unsupported("Native comment has multiple or duplicate anchors");
    byNativeId.set(anchor.anchor.id, anchor);
  }
  const blocks = new Map(
    [...document.blocks, ...document.footnotes].map((block) => [
      block.id,
      block,
    ]),
  );
  return [...nativeComments.keys()]
    .sort((left, right) => Number(left) - Number(right))
    .map((nativeId) => {
      const candidate = byNativeId.get(nativeId);
      if (!candidate || !candidate.paragraph.bookmark)
        unsupported("Native comment does not anchor to an agent-docx block");
      const safeCandidate = candidate as (typeof anchors)[number];
      const block = blocks.get(safeCandidate.paragraph.bookmark as BlockId);
      if (!block) unsupported("Native comment anchors an unknown block");
      const safeBlock = block as
        | LegalDocument["blocks"][number]
        | LegalDocument["footnotes"][number];
      const visible = visibleTextForBlock(safeBlock);
      if (
        safeCandidate.paragraph.text !== visible ||
        safeCandidate.anchor.start > safeCandidate.anchor.end ||
        safeCandidate.anchor.end > visible.length ||
        !isCodePointBoundary(visible, safeCandidate.anchor.start) ||
        !isCodePointBoundary(visible, safeCandidate.anchor.end)
      )
        unsupported("Native comment range is not code-point safe");
      const digest = createHash("sha256")
        .update(`${safeCandidate.paragraph.bookmark}:${nativeId}`)
        .digest("hex");
      const annotationId =
        `a_${digest.slice(0, 8)}-${digest.slice(8, 12)}-${digest.slice(
          12,
          16,
        )}-${digest.slice(16, 20)}-${digest.slice(20, 32)}` as `a_${string}`;
      const native = nativeComments.get(nativeId)!;
      const blockWide =
        safeCandidate.anchor.start === 0 &&
        safeCandidate.anchor.end === visible.length;
      return {
        id: annotationId,
        blockId: safeBlock.id,
        ...(blockWide
          ? {}
          : {
              start: safeCandidate.anchor.start,
              end: safeCandidate.anchor.end,
            }),
        author: native.author === "" ? null : { name: native.author },
        createdAt: native.date,
        message: native.text,
        status: "open" as const,
      };
    });
};
export const parseNativeComments = (
  xml: string | undefined,
): ReadonlyMap<string, NativeComment> => {
  const comments = new Map<string, NativeComment>();
  if (!xml) return comments;
  let current: NativeComment | null = null;
  let inText = false;
  let paragraphCount = 0;
  parseDocxXml(
    xml,
    (tag) => {
      if (tag.local === "comment") {
        if (current) unsupported("Nested DOCX comments are unsupported");
        const id = docxXmlAttribute(tag, "id");
        const author = docxXmlAttribute(tag, "author");
        const date = docxXmlAttribute(tag, "date");
        if (
          !id ||
          !/^\d+$/.test(id) ||
          author === undefined ||
          comments.has(id)
        )
          unsupported("DOCX comment has invalid native attribution");
        if (date !== undefined && Number.isNaN(new Date(date).valueOf()))
          unsupported("DOCX comment has an invalid native date");
        current = {
          id: id as string,
          author: author as string,
          date: date ?? null,
          text: "",
        };
        paragraphCount = 0;
        return;
      }
      if (!current) return;
      if (tag.local === "p") {
        if (paragraphCount > 0) current.text += "\n";
        paragraphCount++;
      } else if (tag.local === "t") inText = true;
    },
    (tag) => {
      if (tag.local === "t") {
        inText = false;
        return;
      }
      if (tag.local === "comment") {
        const closedComment = current;
        if (!closedComment) unsupported("DOCX comment nesting is malformed");
        comments.set(
          (closedComment as NativeComment).id,
          closedComment as NativeComment,
        );
        current = null;
      }
    },
    (text) => {
      if (current && inText) current.text += text;
    },
  );
  if (current || inText) unsupported("DOCX comment text nesting is malformed");
  return comments;
};