import { createHash } from "node:crypto";
import canonicalize from "canonicalize";
import { AgentDocxError, type JsonValue } from "../types.js";
import type {
  Actor,
  AddressableBlock,
  BlockId,
  InlineRun,
  LegalBlock,
  LegalDocument,
  ReviewAnnotation,
  RevisionId,
} from "../legal/model.js";
import type {
  AnnotationChange,
  AttributionSpan,
  BlockLocation,
  Change,
  ChangeAttribution,
  ChangeSet,
  ContainerShell,
  RevisionDeltaRecord,
} from "./types.js";
type FlatBlock = {
  block: AddressableBlock;
  location: BlockLocation;
};

const canonicalHash = (value: unknown): string => {
  const serialized = canonicalize(value);
  if (serialized === undefined) throw new Error("Cannot canonicalize change");
  return createHash("sha256").update(serialized).digest("hex");
};

const withoutPositions = (value: unknown): unknown => {
  if (Array.isArray(value)) return value.map(withoutPositions);
  if (value !== null && typeof value === "object")
    return Object.fromEntries(
      Object.entries(value as Record<string, unknown>)
        .filter(([key]) => key !== "position")
        .map(([key, entry]) => [key, withoutPositions(entry)]),
    );
  return value;
};

const canonicalBlock = (block: AddressableBlock): string => {
  const serialized = canonicalize(withoutPositions(block));
  if (serialized === undefined) throw new Error("Cannot canonicalize block");
  return serialized;
};

const withChangeId = <Value extends Record<string, unknown>>(
  value: Value,
): Value & { id: `c_${string}` } => ({
  ...value,
  id: `c_${canonicalHash(value)}`,
});

const spansFor = (text: string, attribution: ChangeAttribution): readonly AttributionSpan[] =>
  text.length === 0 ? [] : [{ start: 0, end: text.length, attribution }];

const visibleRuns = (runs: readonly InlineRun[]): string =>
  runs.map((run) => `${run.text}${run.hardBreakAfter ? "\n" : ""}`).join("");

const visibleBlock = (block: AddressableBlock): string => {
  if (block.kind === "footnote")
    return block.paragraphs.map((paragraph) => visibleRuns(paragraph.runs)).join("\n");
  if (
    block.kind === "paragraph" ||
    block.kind === "blockquote" ||
    block.kind === "heading" ||
    block.kind === "numbered-paragraph"
  )
    return visibleRuns(block.runs);
  if (block.kind === "list")
    return block.items
      .flatMap((item) => [
        ...item.paragraphs.map((paragraph) => visibleRuns(paragraph.runs)),
        ...item.children.map((child) => visibleBlock(child)),
      ])
      .join("\n");
  if (block.kind === "table")
    return block.rows
      .map((row) =>
        row
          .map((cell) =>
            cell.paragraphs
              .map((paragraph) => visibleRuns(paragraph.runs))
              .join("\n"),
          )
          .join("\t"),
      )
      .join("\n");
  if (block.kind === "exhibit" || block.kind === "length-exclusion")
    return block.blocks.map(visibleBlock).join("\n");
  if (block.kind === "image") return block.alt;
  if (block.kind === "signature") return block.counselId;
  if (block.kind === "certificate") return block.certificateId;
  return "";
};

export const visibleTextForBlock = (block: AddressableBlock): string =>
  visibleBlock(block);

const flatten = (
  blocks: readonly LegalBlock[],
  collection: "body" | "footnotes",
  parentId: BlockId | null,
  result: FlatBlock[],
) => {
  for (const [index, block] of blocks.entries()) {
    result.push({
      block,
      location: {
        collection,
        parentId,
        index,
        sourceOffset: block.position.start.offset,
      },
    });
    if (block.kind === "exhibit" || block.kind === "length-exclusion")
      flatten(block.blocks, collection, block.id, result);
    if (block.kind === "list") {
      for (const item of block.items)
        flatten(item.children, collection, block.id, result);
    }
  }
};

const flattenDocument = (document: LegalDocument): FlatBlock[] => {
  const result: FlatBlock[] = [];
  flatten(document.blocks, "body", null, result);
  for (const [index, footnote] of document.footnotes.entries())
    result.push({
      block: footnote,
      location: {
        collection: "footnotes",
        parentId: null,
        index,
        sourceOffset: footnote.position.start.offset,
      },
    });
  return result;
};

const sourceRange = (block: AddressableBlock) => ({
  start: block.position.start.offset,
  end: block.position.end.offset,
  text: block.sourceText,
});

const shellFor = (block: Extract<LegalBlock, { kind: "list" | "exhibit" | "length-exclusion" }>): ContainerShell => ({
  blockId: block.id,
  kind: block.kind,
  attributes: { kind: block.kind } as JsonValue,
  sourceRanges: [sourceRange(block)],
});

const isContainer = (
  block: AddressableBlock,
): block is Extract<LegalBlock, { kind: "list" | "exhibit" | "length-exclusion" }> =>
  block.kind === "list" || block.kind === "exhibit" || block.kind === "length-exclusion";

const annotationChanges = (
  base: readonly ReviewAnnotation[],
  head: readonly ReviewAnnotation[],
): AnnotationChange[] => {
  const baseById = new Map(base.map((annotation) => [annotation.id, annotation]));
  const headById = new Map(head.map((annotation) => [annotation.id, annotation]));
  const changes: AnnotationChange[] = [];
  for (const id of [...new Set([...baseById.keys(), ...headById.keys()])].sort()) {
    const oldValue = baseById.get(id);
    const newValue = headById.get(id);
    if (!oldValue && newValue) changes.push(withChangeId({ kind: "add", newValue }));
    else if (oldValue && !newValue)
      changes.push(withChangeId({ kind: "remove", oldValue }));
    else if (
      oldValue &&
      newValue &&
      canonicalize(oldValue) !== canonicalize(newValue)
    )
      changes.push(withChangeId({ kind: "replace", oldValue, newValue }));
  }
  return changes.sort((left, right) => left.id.localeCompare(right.id));
};

export const createChangeSet = (
  documentId: string,
  baseRevision: RevisionId,
  headRevision: RevisionId,
  base: LegalDocument,
  head: LegalDocument,
  baseAnnotations: readonly ReviewAnnotation[],
  headAnnotations: readonly ReviewAnnotation[],
  attribution: ChangeAttribution,
): ChangeSet => {
  const baseById = new Map(flattenDocument(base).map((entry) => [entry.block.id, entry]));
  const headById = new Map(flattenDocument(head).map((entry) => [entry.block.id, entry]));
  const changes: Change[] = [];
  const allIds = [...new Set([...baseById.keys(), ...headById.keys()])].sort();
  for (const id of allIds) {
    const previous = baseById.get(id);
    const next = headById.get(id);
    if (!previous && next) {
      const text = visibleBlock(next.block);
      changes.push(
        withChangeId({
          kind: "insert-block",
          blockId: next.block.id,
          to: next.location,
          newSource: sourceRange(next.block),
          block: next.block,
          newAttributionSpans: spansFor(text, attribution),
          attribution,
        }),
      );
      continue;
    }
    if (previous && !next) {
      const text = visibleBlock(previous.block);
      changes.push(
        withChangeId({
          kind: "delete-block",
          blockId: previous.block.id,
          from: previous.location,
          oldSource: sourceRange(previous.block),
          oldBlock: previous.block,
          oldAttributionSpans: spansFor(text, attribution),
          attribution,
        }),
      );
      continue;
    }
    if (!previous || !next) continue;
    const moved =
      previous.location.collection !== next.location.collection ||
      previous.location.parentId !== next.location.parentId ||
      previous.location.index !== next.location.index;
    const oldText = visibleBlock(previous.block);
    const newText = visibleBlock(next.block);
    if (moved) {
      changes.push(
        withChangeId({
          kind: "move-block",
          blockId: next.block.id,
          from: previous.location,
          to: next.location,
          oldSource: sourceRange(previous.block),
          newSource: sourceRange(next.block),
          oldAttributionSpans: spansFor(oldText, attribution),
          newAttributionSpans: spansFor(newText, attribution),
          attribution,
        }),
      );
      continue;
    }
    if (canonicalBlock(previous.block) === canonicalBlock(next.block)) continue;
    if (isContainer(previous.block) && isContainer(next.block)) {
      changes.push(
        withChangeId({
          kind: "replace-container-shell",
          blockId: next.block.id,
          oldShell: shellFor(previous.block),
          newShell: shellFor(next.block),
          oldAttributionSpans: spansFor(oldText, attribution),
          newAttributionSpans: spansFor(newText, attribution),
          attribution,
        }),
      );
      continue;
    }
    if (previous.block.kind === next.block.kind && oldText !== newText) {
      changes.push(
        withChangeId({
          kind: "replace-text",
          blockId: next.block.id,
          oldSource: sourceRange(previous.block),
          newSource: sourceRange(next.block),
          oldText,
          newText,
          oldAttributionSpans: spansFor(oldText, attribution),
          newAttributionSpans: spansFor(newText, attribution),
          attribution,
        }),
      );
      continue;
    }
    if (!isContainer(previous.block) && !isContainer(next.block)) {
      changes.push(
        withChangeId({
          kind: "replace-block",
          blockId: next.block.id,
          oldBlock: previous.block,
          newBlock: next.block,
          oldAttributionSpans: spansFor(oldText, attribution),
          newAttributionSpans: spansFor(newText, attribution),
          attribution,
        }),
      );
    }
  }
  const ordered = changes.sort((left, right) => {
    const leftOffset =
      "from" in left
        ? left.from.sourceOffset
        : "to" in left
          ? left.to.sourceOffset
          : "oldSource" in left
            ? left.oldSource.start
            : "newSource" in left
              ? left.newSource.start
              : Number.MAX_SAFE_INTEGER;
    const rightOffset =
      "from" in right
        ? right.from.sourceOffset
        : "to" in right
          ? right.to.sourceOffset
          : "oldSource" in right
            ? right.oldSource.start
            : "newSource" in right
              ? right.newSource.start
              : Number.MAX_SAFE_INTEGER;
    return leftOffset - rightOffset || left.id.localeCompare(right.id);
  });
  const annotations = annotationChanges(baseAnnotations, headAnnotations);
  const withoutId = {
    schemaVersion: 1 as const,
    documentId,
    baseRevision,
    headRevision,
    changes: ordered,
    annotations,
  };
  return { ...withoutId, id: `sha256:${canonicalHash(withoutId)}` };
};

export const createRevisionDelta = (
  parent: {
    id: RevisionId;
    sourceObject: RevisionId;
    documentConfigObject: RevisionId;
  },
  documentId: string,
  base: LegalDocument,
  head: LegalDocument,
  baseAnnotations: readonly ReviewAnnotation[],
  headAnnotations: readonly ReviewAnnotation[],
  attribution: ChangeAttribution,
): RevisionDeltaRecord => {
  const changeSet = createChangeSet(
    documentId,
    parent.id,
    parent.id,
    base,
    head,
    baseAnnotations,
    headAnnotations,
    attribution,
  );
  return {
    schemaVersion: 1,
    parentSourceObject: parent.sourceObject,
    parentDocumentConfigObject: parent.documentConfigObject,
    changes: changeSet.changes,
    annotations: changeSet.annotations,
  };
};

export const defaultAttribution = (actor: Actor, createdAt: string): ChangeAttribution => ({
  author: actor,
  createdAt,
});
export const rebaseOpenAnnotations = (
  base: LegalDocument,
  head: LegalDocument,
  annotations: readonly ReviewAnnotation[],
): readonly ReviewAnnotation[] => {
  const baseById = new Map(flattenDocument(base).map((entry) => [entry.block.id, entry.block]));
  const headById = new Map(flattenDocument(head).map((entry) => [entry.block.id, entry.block]));
  return annotations.map((annotation) => {
    if (annotation.status !== "open") return annotation;
    const previous = baseById.get(annotation.blockId);
    const next = headById.get(annotation.blockId);
    if (!previous || !next)
      throw new AgentDocxError("ANNOTATION_CONFLICT", `Annotation block was deleted: ${annotation.id}`);
    if (!annotation.range) return annotation;
    const oldText = visibleBlock(previous);
    const newText = visibleBlock(next);
    const { start, end } = annotation.range;
    const boundary = (text: string, offset: number): boolean =>
      Number.isInteger(offset) &&
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
    if (!boundary(oldText, start) || !boundary(oldText, end) || start > end)
      throw new AgentDocxError("ANNOTATION_CONFLICT", `Annotation range is invalid: ${annotation.id}`);
    let prefix = 0;
    while (
      prefix < oldText.length &&
      prefix < newText.length &&
      oldText[prefix] === newText[prefix]
    )
      prefix++;
    let suffix = 0;
    while (
      suffix < oldText.length - prefix &&
      suffix < newText.length - prefix &&
      oldText[oldText.length - suffix - 1] === newText[newText.length - suffix - 1]
    )
      suffix++;
    const oldChangedEnd = oldText.length - suffix;
    const newChangedEnd = newText.length - suffix;
    const translate = (offset: number, affinity: "right" | "left"): number => {
      if (offset < prefix) return offset;
      if (offset > oldChangedEnd)
        return offset + (newChangedEnd - oldChangedEnd);
      if (offset === prefix && prefix === oldChangedEnd)
        return affinity === "right" ? newChangedEnd : prefix;
      if (offset === prefix && oldChangedEnd === newChangedEnd) return offset;
      if (offset === oldChangedEnd) return newChangedEnd;
      throw new AgentDocxError(
        "ANNOTATION_CONFLICT",
        `Annotation range overlaps an edit: ${annotation.id}`,
      );
    };
    const range = {
      start: translate(start, "right"),
      end: translate(end, "left"),
    };
    if (
      !boundary(newText, range.start) ||
      !boundary(newText, range.end) ||
      range.start > range.end
    )
      throw new AgentDocxError(
        "ANNOTATION_CONFLICT",
        `Annotation range cannot be rebased: ${annotation.id}`,
      );
    return { ...annotation, range };
  });
};
