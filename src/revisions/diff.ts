import { createHash } from "node:crypto";
import canonicalize from "canonicalize";
import { AgentDocxError, type JsonValue } from "../types.js";
import type {
  Actor,
  AddressableBlock,
  BlockId,
  LegalBlock,
  LegalDocument,
  ReviewAnnotation,
  RevisionId,
} from "../legal/model.js";
import { visibleBlock } from "../legal/visible-text.js";
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
export type JsonObject = { readonly [key: string]: JsonValue };
const compareText = (left: string, right: string): number =>
  left < right ? -1 : left > right ? 1 : 0;

export type ChangeSetOptions = {
  baseConfig?: JsonObject;
  headConfig?: JsonObject;
  baseDependencies?: Readonly<Record<string, RevisionId>>;
  headDependencies?: Readonly<Record<string, RevisionId>>;
  baseSource?: string;
  headSource?: string;
};

const jsonObject = (value: unknown): JsonObject => value as JsonObject;

const pointerSegment = (value: string): string =>
  value.replaceAll("~", "~0").replaceAll("/", "~1");

type PresentValue = { present: false } | { present: true; value: JsonValue };

const configChanges = (
  base: PresentValue,
  head: PresentValue,
  path: string,
  attribution: ChangeAttribution,
  changes: Change[],
): void => {
  if (!base.present && !head.present) return;
  if (!base.present) {
    if (!head.present) return;
    changes.push(
      withChangeId({
        kind: "add-config",
        path,
        newValue: head.value,
        attribution,
      }),
    );
    return;
  }
  if (!head.present) {
    changes.push(
      withChangeId({
        kind: "remove-config",
        path,
        oldValue: base.value,
        attribution,
      }),
    );
    return;
  }
  if (
    typeof base.value === "object" &&
    base.value !== null &&
    !Array.isArray(base.value) &&
    typeof head.value === "object" &&
    head.value !== null &&
    !Array.isArray(head.value)
  ) {
    const baseObject = base.value as JsonObject;
    const headObject = head.value as JsonObject;
    const keys = [
      ...new Set([...Object.keys(baseObject), ...Object.keys(headObject)]),
    ].sort(compareText);
    for (const key of keys) {
      configChanges(
        Object.hasOwn(baseObject, key)
          ? { present: true, value: baseObject[key]! }
          : { present: false },
        Object.hasOwn(headObject, key)
          ? { present: true, value: headObject[key]! }
          : { present: false },
        `${path}/${pointerSegment(key)}`,
        attribution,
        changes,
      );
    }
    return;
  }
  if (canonicalize(base.value) === canonicalize(head.value)) return;
  changes.push(
    withChangeId({
      kind: "replace-config",
      path,
      oldValue: base.value,
      newValue: head.value,
      attribution,
    }),
  );
};

const dependencyChanges = (
  base: Readonly<Record<string, RevisionId>>,
  head: Readonly<Record<string, RevisionId>>,
  attribution: ChangeAttribution,
): Change[] => {
  const changes: Change[] = [];
  const keys = [...new Set([...Object.keys(base), ...Object.keys(head)])].sort(
    compareText,
  );
  for (const key of keys) {
    const oldObject = base[key];
    const newObject = head[key];
    if (oldObject === undefined && newObject !== undefined) {
      changes.push(
        withChangeId({
          kind: "add-dependency",
          key,
          newObject,
          attribution,
        }),
      );
    } else if (oldObject !== undefined && newObject === undefined) {
      changes.push(
        withChangeId({
          kind: "remove-dependency",
          key,
          oldObject,
          attribution,
        }),
      );
    } else if (
      oldObject !== undefined &&
      newObject !== undefined &&
      oldObject !== newObject
    ) {
      changes.push(
        withChangeId({
          kind: "replace-dependency",
          key,
          oldObject,
          newObject,
          attribution,
        }),
      );
    }
  }
  return changes;
};

const configAndDependencyChanges = (
  options: ChangeSetOptions,
  attribution: ChangeAttribution,
): Change[] => {
  const changes: Change[] = [];
  if (options.baseConfig && options.headConfig) {
    configChanges(
      { present: true, value: jsonObject(options.baseConfig) },
      { present: true, value: jsonObject(options.headConfig) },
      "",
      attribution,
      changes,
    );
  }
  if (options.baseDependencies && options.headDependencies)
    changes.push(
      ...dependencyChanges(
        options.baseDependencies,
        options.headDependencies,
        attribution,
      ),
    );
  return changes;
};

type FlatBlock = {
  block: AddressableBlock;
  location: BlockLocation;
};

const canonicalHash = (value: unknown): string => {
  const serialized = canonicalize(value);
  if (serialized === undefined)
    throw new AgentDocxError("INTERNAL_ERROR", "Cannot canonicalize change");
  return createHash("sha256").update(serialized).digest("hex");
};

const withoutPositions = (value: unknown): unknown => {
  if (Array.isArray(value)) return value.map(withoutPositions);
  if (value !== null && typeof value === "object")
    return Object.fromEntries(
      Object.entries(value as Record<string, unknown>)
        .filter(([key]) => key !== "position" && key !== "sourceStartOffset")
        .map(([key, entry]) => [key, withoutPositions(entry)]),
    );
  return value;
};

const canonicalBlock = (block: AddressableBlock): string => {
  const serialized = canonicalize(withoutPositions(block));
  if (serialized === undefined)
    throw new AgentDocxError("INTERNAL_ERROR", "Cannot canonicalize block");
  return serialized;
};

const withChangeId = <const Value extends Record<string, unknown>>(
  value: Value,
): Omit<Value, "id"> & { id: `c_${string}` } => {
  const { id: _ignored, ...withoutId } = value as Value & {
    id?: unknown;
  };
  return {
    ...withoutId,
    id: `c_${canonicalHash(withoutId)}`,
  } as Omit<Value, "id"> & { id: `c_${string}` };
};

const spansFor = (
  text: string,
  attribution: ChangeAttribution,
): readonly AttributionSpan[] =>
  text.length === 0 ? [] : [{ start: 0, end: text.length, attribution }];
type TextToken = { start: number; end: number; text: string };

const tokenizeVisibleText = (text: string): readonly TextToken[] => {
  const tokens: TextToken[] = [];
  const pattern =
    /(\s+|\p{L}[\p{L}\p{M}\p{N}'’.-]*|\p{N}+(?:[.,]\p{N}+)*|[^\s])/gu;
  let cursor = 0;
  for (const match of text.matchAll(pattern)) {
    const start = match.index ?? cursor;
    if (start > cursor)
      tokens.push({
        start: cursor,
        end: start,
        text: text.slice(cursor, start),
      });
    const value = match[0]!;
    tokens.push({ start, end: start + value.length, text: value });
    cursor = start + value.length;
  }
  if (cursor < text.length)
    tokens.push({ start: cursor, end: text.length, text: text.slice(cursor) });
  return tokens;
};
const codePointTokens = (text: string): readonly TextToken[] => {
  const tokens: TextToken[] = [];
  let offset = 0;
  for (const value of text) {
    tokens.push({
      start: offset,
      end: offset + value.length,
      text: value,
    });
    offset += value.length;
  }
  return tokens;
};
const MAX_DIFF_TOKENS = 50_000;
const MAX_DIFF_TRACE_CELLS = 4_000_000;

const enforceTokenBudget = (
  oldTokens: readonly TextToken[],
  newTokens: readonly TextToken[],
): void => {
  if (oldTokens.length + newTokens.length > MAX_DIFF_TOKENS)
    throw new AgentDocxError(
      "DIFF_TOO_LARGE",
      `Revision diff exceeds the ${MAX_DIFF_TOKENS}-token budget`,
    );
};

const equalTokenPairs = (
  oldTokens: readonly TextToken[],
  newTokens: readonly TextToken[],
): ReadonlyMap<number, number> => {
  enforceTokenBudget(oldTokens, newTokens);
  const n = oldTokens.length;
  const m = newTokens.length;
  const max = n + m;
  const trace: Map<number, number>[] = [];
  let traceCells = 0;
  let frontier = new Map<number, number>([[1, 0]]);
  let finalDepth = 0;
  for (let depth = 0; depth <= max; depth++) {
    traceCells += frontier.size;
    if (traceCells > MAX_DIFF_TRACE_CELLS)
      throw new AgentDocxError(
        "DIFF_TOO_LARGE",
        `Revision diff exceeds the ${MAX_DIFF_TRACE_CELLS}-cell trace budget`,
      );
    trace.push(new Map(frontier));
    for (let diagonal = -depth; diagonal <= depth; diagonal += 2) {
      const down = frontier.get(diagonal + 1) ?? Number.NEGATIVE_INFINITY;
      const right = frontier.get(diagonal - 1) ?? Number.NEGATIVE_INFINITY;
      let x: number;
      if (diagonal === -depth || (diagonal !== depth && right < down)) x = down;
      else x = right + 1;
      if (!Number.isFinite(x)) continue;
      let y = x - diagonal;
      while (x < n && y < m && oldTokens[x]!.text === newTokens[y]!.text) {
        x++;
        y++;
      }
      frontier.set(diagonal, x);
      if (x >= n && y >= m) {
        finalDepth = depth;
        const pairs: [number, number][] = [];
        let currentX = n;
        let currentY = m;
        for (let d = finalDepth; d > 0; d--) {
          const previous = trace[d]!;
          const k = currentX - currentY;
          const downValue = previous.get(k + 1) ?? Number.NEGATIVE_INFINITY;
          const rightValue = previous.get(k - 1) ?? Number.NEGATIVE_INFINITY;
          const previousDiagonal =
            k === -d || (k !== d && rightValue < downValue) ? k + 1 : k - 1;
          const previousX = previous.get(previousDiagonal) ?? 0;
          const previousY = previousX - previousDiagonal;
          while (currentX > previousX && currentY > previousY) {
            pairs.push([currentX - 1, currentY - 1]);
            currentX--;
            currentY--;
          }
          currentX = previousX;
          currentY = previousY;
        }
        while (
          currentX > 0 &&
          currentY > 0 &&
          oldTokens[currentX - 1]!.text === newTokens[currentY - 1]!.text
        ) {
          pairs.push([currentX - 1, currentY - 1]);
          currentX--;
          currentY--;
        }
        return new Map(
          pairs.map(([oldIndex, newIndex]) => [newIndex, oldIndex]),
        );
      }
    }
  }
  return new Map();
};

const sameAttribution = (
  left: ChangeAttribution,
  right: ChangeAttribution,
): boolean => canonicalize(left) === canonicalize(right);

const appendAttributionSpan = (
  spans: AttributionSpan[],
  start: number,
  end: number,
  attribution: ChangeAttribution,
): void => {
  if (end <= start) return;
  const previous = spans.at(-1);
  if (
    previous &&
    previous.end === start &&
    sameAttribution(previous.attribution, attribution)
  ) {
    spans[spans.length - 1] = { ...previous, end };
  } else spans.push({ start, end, attribution });
};

export const reattributeVisibleText = (
  oldText: string,
  newText: string,
  oldSpans: readonly AttributionSpan[] | undefined,
  insertedAttribution: ChangeAttribution,
): readonly AttributionSpan[] => {
  if (newText.length === 0) return [];
  if (!oldSpans || oldText.length === 0)
    return spansFor(newText, insertedAttribution);
  const oldTokens = tokenizeVisibleText(oldText);
  const newTokens = tokenizeVisibleText(newText);
  const matches = equalTokenPairs(oldTokens, newTokens);
  const spans: AttributionSpan[] = [];
  for (let newIndex = 0; newIndex < newTokens.length; newIndex++) {
    const newToken = newTokens[newIndex]!;
    const oldIndex = matches.get(newIndex);
    if (oldIndex === undefined) {
      appendAttributionSpan(
        spans,
        newToken.start,
        newToken.end,
        insertedAttribution,
      );
      continue;
    }
    const oldToken = oldTokens[oldIndex]!;
    let cursor = oldToken.start;
    for (const span of oldSpans) {
      if (span.end <= oldToken.start) continue;
      if (span.start >= oldToken.end) break;
      const start = Math.max(span.start, oldToken.start);
      const end = Math.min(span.end, oldToken.end);
      if (start > cursor)
        appendAttributionSpan(
          spans,
          newToken.start + cursor - oldToken.start,
          newToken.start + start - oldToken.start,
          insertedAttribution,
        );
      appendAttributionSpan(
        spans,
        newToken.start + start - oldToken.start,
        newToken.start + end - oldToken.start,
        span.attribution,
      );
      cursor = end;
    }
    if (cursor < oldToken.end)
      appendAttributionSpan(
        spans,
        newToken.start + cursor - oldToken.start,
        newToken.end,
        insertedAttribution,
      );
  }
  return spans.length > 0 ? spans : spansFor(newText, insertedAttribution);
};

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

const blockSourceRange = (block: AddressableBlock) => ({
  start: block.position.start.offset,
  end: block.position.end.offset,
  text: block.sourceText,
});

const structuralSourceRange = (block: AddressableBlock, source?: string) => {
  if (source === undefined) return blockSourceRange(block);
  const start = block.position.start.offset;
  const contentLineStart = source.lastIndexOf("\n", start - 1) + 1;
  const markerLineEnd = contentLineStart - 1;
  const markerContentEnd =
    markerLineEnd >= 0 && source[markerLineEnd - 1] === "\r"
      ? markerLineEnd - 1
      : markerLineEnd;
  const markerLineStart =
    markerContentEnd >= 0
      ? source.lastIndexOf("\n", markerContentEnd - 1) + 1
      : 0;
  const markerLine = source.slice(markerLineStart, markerContentEnd);
  const markerId = markerLine.match(
    /^[ \t]*<!--[ \t]*agent-docx:block[ \t]+id="([^"]+)"[ \t]*-->$/,
  )?.[1];
  if (markerId === block.id)
    return {
      start: markerLineStart,
      end: block.position.end.offset,
      text: source.slice(markerLineStart, block.position.end.offset),
    };
  return blockSourceRange(block);
};
const shellSourceRanges = (
  block: Extract<LegalBlock, { kind: "list" | "exhibit" | "length-exclusion" }>,
  source: string | undefined,
): readonly { start: number; end: number; text: string }[] => {
  const whole = structuralSourceRange(block, source);
  if (source === undefined) return [whole];
  const children =
    block.kind === "list"
      ? block.items.flatMap((item) => item.children)
      : block.blocks;
  if (children.length === 0) return [whole];
  const childRanges = children.map((child) => {
    const range = structuralSourceRange(child, source);
    return { start: range.start, end: range.end };
  });
  const ranges = childRanges
    .filter((range) => range.start >= whole.start && range.end <= whole.end)
    .sort((left, right) => left.start - right.start || left.end - right.end);
  const result: { start: number; end: number }[] = [];
  let cursor = whole.start;
  for (const range of ranges) {
    if (range.start < cursor) {
      cursor = Math.max(cursor, range.end);
      continue;
    }
    if (cursor < range.start)
      result.push({
        start: cursor,
        end: range.start,
      });
    cursor = range.end;
  }
  if (cursor < whole.end)
    result.push({
      start: cursor,
      end: whole.end,
    });
  return result.length > 0
    ? result.map((range) => ({
        ...range,
        text: source.slice(range.start, range.end),
      }))
    : [whole];
};
const shellFor = (
  block: Extract<LegalBlock, { kind: "list" | "exhibit" | "length-exclusion" }>,
  source?: string,
): ContainerShell => ({
  blockId: block.id,
  kind: block.kind,
  attributes:
    block.kind === "list"
      ? { ordered: block.ordered, start: block.start, depth: block.depth }
      : block.kind === "exhibit"
        ? {
            exhibitId: block.exhibitId,
            label: block.label,
            source: block.source,
          }
        : {
            exclusionKind: block.exclusionKind,
            ...(block.citation ? { citation: block.citation } : {}),
          },
  sourceRanges: shellSourceRanges(block, source),
});
const containerShellSignature = (
  block: Extract<LegalBlock, { kind: "list" | "exhibit" | "length-exclusion" }>,
  source?: string,
): string => {
  const shell = shellFor(block, source);
  const serialized = canonicalize({
    kind: shell.kind,
    attributes: shell.attributes,
    sourceTexts: shell.sourceRanges.map((range) => range.text),
  });
  if (serialized === undefined)
    throw new AgentDocxError(
      "INTERNAL_ERROR",
      "Cannot canonicalize container shell",
    );
  return serialized;
};

const isContainer = (
  block: AddressableBlock,
): block is Extract<
  LegalBlock,
  { kind: "list" | "exhibit" | "length-exclusion" }
> =>
  block.kind === "list" ||
  block.kind === "exhibit" ||
  block.kind === "length-exclusion";
type TextLeaf = Extract<
  AddressableBlock,
  { kind: "paragraph" | "blockquote" | "heading" | "numbered-paragraph" }
>;

const isTextLeaf = (block: AddressableBlock): block is TextLeaf =>
  block.kind === "paragraph" ||
  block.kind === "blockquote" ||
  block.kind === "heading" ||
  block.kind === "numbered-paragraph";

type ExactTextMapping = {
  map: (offset: number) => number | null;
};

const exactTextMapping = (
  block: TextLeaf,
  text: string,
): ExactTextMapping | null => {
  if (block.segments.length === 0) {
    const sourceStart = block.position.start.offset;
    if (
      block.sourceText !== text ||
      block.position.end.offset - sourceStart !== text.length
    )
      return null;
    return {
      map: (offset: number): number | null =>
        Number.isInteger(offset) && offset >= 0 && offset <= text.length
          ? sourceStart + offset
          : null,
    };
  }
  let normalizedEnd = 0;
  let sourceEnd: number | undefined;
  for (const segment of block.segments) {
    if (segment.normalizedStart !== normalizedEnd) return null;
    const width = segment.normalizedEnd - segment.normalizedStart;
    const sourceWidth =
      segment.position.end.offset - segment.position.start.offset;
    if (
      width < 0 ||
      sourceWidth < 0 ||
      segment.sourceStartOffset !== segment.position.start.offset ||
      (segment.precision === "exact" && sourceWidth !== width) ||
      (sourceEnd !== undefined && segment.sourceStartOffset !== sourceEnd)
    )
      return null;
    normalizedEnd = segment.normalizedEnd;
    sourceEnd = segment.position.end.offset;
  }
  if (normalizedEnd !== text.length) return null;
  const first = block.segments[0];
  const sourceStart = first?.sourceStartOffset ?? block.position.start.offset;
  const finalSourceEnd = sourceEnd ?? sourceStart;
  return {
    map: (offset: number): number | null => {
      if (!Number.isInteger(offset) || offset < 0 || offset > text.length)
        return null;
      if (offset === text.length) return finalSourceEnd;
      const segment = block.segments.find(
        (entry) =>
          entry.normalizedStart <= offset && offset < entry.normalizedEnd,
      );
      if (!segment) return null;
      if (offset === segment.normalizedStart)
        return segment.position.start.offset;
      if (offset === segment.normalizedEnd) return segment.position.end.offset;
      return segment.sourceStartOffset + (offset - segment.normalizedStart);
    },
  };
};

const sourceSliceFor = (
  block: TextLeaf,
  source: string | undefined,
  start: number,
  end: number,
  mapping: ExactTextMapping,
): { start: number; end: number; text: string } | null => {
  const sourceStart = mapping.map(start);
  const sourceEnd = mapping.map(end);
  if (sourceStart === null || sourceEnd === null || sourceEnd < sourceStart)
    return null;
  if (source !== undefined)
    return {
      start: sourceStart,
      end: sourceEnd,
      text: source.slice(sourceStart, sourceEnd),
    };
  const localStart = sourceStart - block.position.start.offset;
  const localEnd = sourceEnd - block.position.start.offset;
  if (localStart < 0 || localEnd > block.sourceText.length) return null;
  return {
    start: sourceStart,
    end: sourceEnd,
    text: block.sourceText.slice(localStart, localEnd),
  };
};

type TextEdit =
  | {
      kind: "insert";
      oldStart: number;
      oldEnd: number;
      newStart: number;
      newEnd: number;
    }
  | {
      kind: "delete";
      oldStart: number;
      oldEnd: number;
      newStart: number;
      newEnd: number;
    }
  | {
      kind: "replace";
      oldStart: number;
      oldEnd: number;
      newStart: number;
      newEnd: number;
    };

const textEdits = (oldText: string, newText: string): readonly TextEdit[] => {
  const oldTokens = tokenizeVisibleText(oldText);
  const newTokens = tokenizeVisibleText(newText);
  const pairs = [...equalTokenPairs(oldTokens, newTokens).entries()].sort(
    ([newLeft, oldLeft], [newRight, oldRight]) =>
      oldLeft - oldRight || newLeft - newRight,
  );
  const edits: TextEdit[] = [];
  let oldCursor = 0;
  let newCursor = 0;
  const pushGap = (oldIndex: number, newIndex: number): void => {
    if (oldIndex === oldCursor && newIndex === newCursor) return;
    const oldStart =
      oldCursor < oldTokens.length
        ? oldTokens[oldCursor]!.start
        : oldText.length;
    const oldEnd =
      oldIndex > oldCursor ? oldTokens[oldIndex - 1]!.end : oldStart;
    const newStart =
      newCursor < newTokens.length
        ? newTokens[newCursor]!.start
        : newText.length;
    const newEnd =
      newIndex > newCursor ? newTokens[newIndex - 1]!.end : newStart;
    edits.push({
      kind:
        oldEnd === oldStart
          ? "insert"
          : newEnd === newStart
            ? "delete"
            : "replace",
      oldStart,
      oldEnd,
      newStart,
      newEnd,
    });
  };
  for (const [newIndex, oldIndex] of pairs) {
    if (oldIndex < oldCursor || newIndex < newCursor) return [];
    pushGap(oldIndex, newIndex);
    oldCursor = oldIndex + 1;
    newCursor = newIndex + 1;
  }
  pushGap(oldTokens.length, newTokens.length);
  const coalesced: TextEdit[] = [];
  for (const edit of edits) {
    const previous = coalesced.at(-1);
    const oldGap = previous && oldText.slice(previous.oldEnd, edit.oldStart);
    const newGap = previous && newText.slice(previous.newEnd, edit.newStart);
    if (
      previous &&
      (previous.kind === "replace" || edit.kind === "replace") &&
      oldGap !== undefined &&
      newGap !== undefined &&
      /^\s*$/.test(oldGap) &&
      /^\s*$/.test(newGap)
    ) {
      previous.oldEnd = edit.oldEnd;
      previous.newEnd = edit.newEnd;
      previous.kind =
        previous.oldEnd === previous.oldStart
          ? "insert"
          : previous.newEnd === previous.newStart
            ? "delete"
            : "replace";
    } else {
      coalesced.push({ ...edit });
    }
  }
  return coalesced;
};

type TextChange = {
  kind: "insert-text" | "delete-text" | "replace-text";
  oldStart: number;
  oldEnd: number;
  newStart: number;
  newEnd: number;
};

type MappedTextChange = TextChange & {
  oldSource: { start: number; end: number; text: string };
  newSource: { start: number; end: number; text: string };
  oldText: string;
  newText: string;
};

const mappedTextChanges = (
  previous: TextLeaf,
  next: TextLeaf,
  oldText: string,
  newText: string,
  baseSource: string | undefined,
  headSource: string | undefined,
): readonly MappedTextChange[] | null => {
  const oldMapping = exactTextMapping(previous, oldText);
  const newMapping = exactTextMapping(next, newText);
  if (!oldMapping || !newMapping) return null;
  const edits = textEdits(oldText, newText);
  if (oldText !== newText && edits.length === 0) return null;
  const result: MappedTextChange[] = [];
  for (const edit of edits) {
    const oldSource = sourceSliceFor(
      previous,
      baseSource,
      edit.oldStart,
      edit.oldEnd,
      oldMapping,
    );
    const newSource = sourceSliceFor(
      next,
      headSource,
      edit.newStart,
      edit.newEnd,
      newMapping,
    );
    if (!oldSource || !newSource) return null;
    result.push({
      ...edit,
      kind:
        edit.kind === "insert"
          ? "insert-text"
          : edit.kind === "delete"
            ? "delete-text"
            : "replace-text",
      oldSource,
      newSource,
      oldText: oldText.slice(edit.oldStart, edit.oldEnd),
      newText: newText.slice(edit.newStart, edit.newEnd),
    });
  }
  return result;
};

const annotationChanges = (
  base: readonly ReviewAnnotation[],
  head: readonly ReviewAnnotation[],
): AnnotationChange[] => {
  const baseById = new Map(
    base.map((annotation) => [annotation.id, annotation]),
  );
  const headById = new Map(
    head.map((annotation) => [annotation.id, annotation]),
  );
  const changes: AnnotationChange[] = [];
  for (const id of [
    ...new Set([...baseById.keys(), ...headById.keys()]),
  ].sort()) {
    const oldValue = baseById.get(id);
    const newValue = headById.get(id);
    if (!oldValue && newValue)
      changes.push(withChangeId({ kind: "add", newValue }));
    else if (oldValue && !newValue)
      changes.push(withChangeId({ kind: "remove", oldValue }));
    else if (
      oldValue &&
      newValue &&
      canonicalize(oldValue) !== canonicalize(newValue)
    )
      changes.push(withChangeId({ kind: "replace", oldValue, newValue }));
  }
  return changes.sort((left, right) => compareText(left.id, right.id));
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
  options: ChangeSetOptions = {},
): ChangeSet => {
  const baseById = new Map(
    flattenDocument(base).map((entry) => [entry.block.id, entry]),
  );
  const headById = new Map(
    flattenDocument(head).map((entry) => [entry.block.id, entry]),
  );
  const changes: Change[] = [];
  const allIds = [...new Set([...baseById.keys(), ...headById.keys()])].sort();
  const hasAncestor = (
    entry: FlatBlock,
    lookup: ReadonlyMap<string, FlatBlock>,
    ancestors: ReadonlySet<string>,
  ): boolean => {
    let parentId = entry.location.parentId;
    while (parentId !== null) {
      if (ancestors.has(parentId)) return true;
      parentId = lookup.get(parentId)?.location.parentId ?? null;
    }
    return false;
  };
  const deletedIds = new Set(
    [...baseById.keys()].filter((id) => !headById.has(id)),
  );
  const insertedIds = new Set(
    [...headById.keys()].filter((id) => !baseById.has(id)),
  );
  const movedContainerIds = new Set(
    [...baseById.keys()].filter((id) => {
      const previous = baseById.get(id);
      const next = headById.get(id);
      return (
        previous !== undefined &&
        next !== undefined &&
        isContainer(previous.block) &&
        (previous.location.collection !== next.location.collection ||
          previous.location.parentId !== next.location.parentId ||
          previous.location.index !== next.location.index)
      );
    }),
  );
  const changedContainerIds = new Set(
    [...baseById.keys()].filter((id) => {
      const previous = baseById.get(id);
      const next = headById.get(id);
      return (
        previous !== undefined &&
        next !== undefined &&
        isContainer(previous.block) &&
        isContainer(next.block) &&
        containerShellSignature(previous.block, options.baseSource) !==
          containerShellSignature(next.block, options.headSource)
      );
    }),
  );
  for (const id of allIds) {
    const previous = baseById.get(id);
    const next = headById.get(id);
    if (!previous && next) {
      if (
        hasAncestor(next, headById, insertedIds) ||
        hasAncestor(next, headById, changedContainerIds)
      )
        continue;
      const text = visibleBlock(next.block, head.metadata);
      changes.push(
        withChangeId({
          kind: "insert-block",
          blockId: next.block.id,
          to: next.location,
          newSource: structuralSourceRange(next.block, options.headSource),
          block: next.block,
          newAttributionSpans: spansFor(text, attribution),
          attribution,
        }),
      );
      continue;
    }
    if (previous && !next) {
      if (
        hasAncestor(previous, baseById, deletedIds) ||
        hasAncestor(previous, baseById, changedContainerIds)
      )
        continue;
      const text = visibleBlock(previous.block, base.metadata);
      changes.push(
        withChangeId({
          kind: "delete-block",
          blockId: previous.block.id,
          from: previous.location,
          oldSource: structuralSourceRange(previous.block, options.baseSource),
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
    const oldText = visibleBlock(previous.block, base.metadata);
    const newText = visibleBlock(next.block, head.metadata);
    if (
      hasAncestor(next, headById, movedContainerIds) ||
      hasAncestor(next, headById, changedContainerIds)
    )
      continue;
    if (moved) {
      changes.push(
        withChangeId({
          kind: "move-block",
          blockId: next.block.id,
          from: previous.location,
          to: next.location,
          oldSource: structuralSourceRange(previous.block, options.baseSource),
          newSource: structuralSourceRange(next.block, options.headSource),
          oldAttributionSpans: spansFor(oldText, attribution),
          newAttributionSpans: spansFor(newText, attribution),
          attribution,
        }),
      );
      continue;
    }
    if (isContainer(previous.block) && isContainer(next.block)) {
      if (
        containerShellSignature(previous.block, options.baseSource) !==
        containerShellSignature(next.block, options.headSource)
      )
        changes.push(
          withChangeId({
            kind: "replace-container-shell",
            blockId: next.block.id,
            oldShell: shellFor(previous.block, options.baseSource),
            newShell: shellFor(next.block, options.headSource),
            oldAttributionSpans: spansFor(oldText, attribution),
            newAttributionSpans: spansFor(newText, attribution),
            attribution,
          }),
        );
      continue;
    }
    if (canonicalBlock(previous.block) === canonicalBlock(next.block)) continue;
    if (
      isTextLeaf(previous.block) &&
      isTextLeaf(next.block) &&
      previous.block.kind === next.block.kind &&
      oldText !== newText
    ) {
      const mapped = mappedTextChanges(
        previous.block,
        next.block,
        oldText,
        newText,
        options.baseSource,
        options.headSource,
      );
      if (mapped) {
        for (const change of mapped) {
          if (change.kind === "insert-text")
            changes.push(
              withChangeId({
                kind: change.kind,
                blockId: next.block.id,
                oldOffset: change.oldSource.start,
                newSource: change.newSource,
                newText: change.newText,
                newAttributionSpans: spansFor(change.newText, attribution),
                attribution,
              }),
            );
          else if (change.kind === "delete-text")
            changes.push(
              withChangeId({
                kind: change.kind,
                blockId: previous.block.id,
                oldSource: change.oldSource,
                newOffset: change.newSource.start,
                oldText: change.oldText,
                oldAttributionSpans: spansFor(change.oldText, attribution),
                attribution,
              }),
            );
          else
            changes.push(
              withChangeId({
                kind: change.kind,
                blockId: next.block.id,
                oldSource: change.oldSource,
                newSource: change.newSource,
                oldText: change.oldText,
                newText: change.newText,
                oldAttributionSpans: spansFor(change.oldText, attribution),
                newAttributionSpans: spansFor(change.newText, attribution),
                attribution,
              }),
            );
        }
        continue;
      }
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
  const kindRank: Readonly<Record<Change["kind"], number>> = {
    "delete-block": 0,
    "move-block": 1,
    "replace-container-shell": 2,
    "replace-block": 3,
    "delete-text": 4,
    "replace-text": 5,
    "insert-text": 6,
    "insert-block": 7,
    "add-config": 8,
    "remove-config": 9,
    "replace-config": 10,
    "add-dependency": 11,
    "remove-dependency": 12,
    "replace-dependency": 13,
  };
  const ancestorChain = (
    location: BlockLocation | undefined,
    lookup: ReadonlyMap<string, FlatBlock>,
  ): string =>
    location
      ? (() => {
          const chain: string[] = [];
          let parentId = location.parentId;
          while (parentId !== null) {
            chain.unshift(parentId);
            parentId = lookup.get(parentId)?.location.parentId ?? null;
          }
          return chain.join("/");
        })()
      : "";
  const sourceMeta = (change: Change) => {
    const blockId = "blockId" in change ? change.blockId : undefined;
    const headEntry = blockId ? headById.get(blockId) : undefined;
    const location =
      change.kind === "delete-block" || change.kind === "move-block"
        ? change.from
        : change.kind === "insert-block"
          ? change.to
          : undefined;
    const oldOffset = (() => {
      switch (change.kind) {
        case "delete-block":
        case "move-block":
        case "delete-text":
        case "replace-text":
          return change.oldSource.start;
        case "insert-text":
          return change.oldOffset;
        case "replace-block":
          return change.oldBlock.position.start.offset;
        case "replace-container-shell":
          return (
            change.oldShell.sourceRanges[0]?.start ?? Number.MAX_SAFE_INTEGER
          );
        default:
          return Number.MAX_SAFE_INTEGER;
      }
    })();
    const newOffset = (() => {
      switch (change.kind) {
        case "insert-block":
        case "move-block":
        case "insert-text":
        case "replace-text":
          return change.newSource.start;
        case "delete-text":
          return change.newOffset;
        case "replace-block":
          return change.newBlock.position.start.offset;
        case "replace-container-shell":
          return (
            change.newShell.sourceRanges[0]?.start ?? Number.MAX_SAFE_INTEGER
          );
        default:
          return Number.MAX_SAFE_INTEGER;
      }
    })();
    const primary = (() => {
      switch (change.kind) {
        case "delete-block":
        case "move-block":
        case "delete-text":
        case "replace-text":
          return change.oldSource.start;
        case "insert-text":
          return change.oldOffset;
        case "replace-block":
          return change.oldBlock.position.start.offset;
        case "replace-container-shell":
          return (
            change.oldShell.sourceRanges[0]?.start ?? Number.MAX_SAFE_INTEGER
          );
        case "insert-block":
          return change.newSource.start;
        default:
          return Number.MAX_SAFE_INTEGER;
      }
    })();
    return {
      collection: location?.collection === "footnotes" ? 1 : 0,
      oldOffset,
      newOffset,
      ancestors: ancestorChain(location, headEntry ? headById : baseById),
      blockId: blockId ?? "",
      kind: kindRank[change.kind],
      primary,
    };
  };
  const orderedSource = changes.sort((left, right) => {
    const leftMeta = sourceMeta(left);
    const rightMeta = sourceMeta(right);
    return (
      leftMeta.collection - rightMeta.collection ||
      leftMeta.oldOffset - rightMeta.oldOffset ||
      leftMeta.newOffset - rightMeta.newOffset ||
      compareText(leftMeta.ancestors, rightMeta.ancestors) ||
      compareText(leftMeta.blockId, rightMeta.blockId) ||
      leftMeta.kind - rightMeta.kind ||
      leftMeta.primary - rightMeta.primary ||
      compareText(left.id, right.id)
    );
  });
  const configAndDependencies = configAndDependencyChanges(
    options,
    attribution,
  );
  const ordered = [
    ...orderedSource,
    ...configAndDependencies
      .filter((change) => change.kind.endsWith("-config"))
      .sort(
        (left, right) =>
          compareText(
            "path" in left ? left.path : "",
            "path" in right ? right.path : "",
          ) ||
          compareText(left.kind, right.kind) ||
          compareText(left.id, right.id),
      ),
    ...configAndDependencies
      .filter((change) => change.kind.endsWith("-dependency"))
      .sort(
        (left, right) =>
          compareText(
            "key" in left ? left.key : "",
            "key" in right ? right.key : "",
          ) ||
          compareText(left.kind, right.kind) ||
          compareText(left.id, right.id),
      ),
  ];
  const annotations = annotationChanges(baseAnnotations, headAnnotations);
  const withoutId = {
    schemaVersion: 1 as const,
    documentId,
    baseRevision,
    headRevision,
    changes: ordered,
    annotations,
  };
  return {
    ...withoutId,
    id: `sha256:${canonicalHash(withoutId)}`,
  };
};

export type ChangeSetProvenance = {
  baseBlocks: ReadonlyMap<string, readonly AttributionSpan[]>;
  headBlocks: ReadonlyMap<string, readonly AttributionSpan[]>;
  baseOperations: ReadonlyMap<string, ChangeAttribution>;
  headOperations: ReadonlyMap<string, ChangeAttribution>;
  baseConfig: ReadonlyMap<string, ChangeAttribution>;
  headConfig: ReadonlyMap<string, ChangeAttribution>;
  baseDependencies: ReadonlyMap<string, ChangeAttribution>;
  headDependencies: ReadonlyMap<string, ChangeAttribution>;
  baseConfigOperations?: ReadonlyMap<string, ChangeAttribution>;
  headConfigOperations?: ReadonlyMap<string, ChangeAttribution>;
  baseDependencyOperations?: ReadonlyMap<string, ChangeAttribution>;
  headDependencyOperations?: ReadonlyMap<string, ChangeAttribution>;
  baseDocument?: LegalDocument;
  headDocument?: LegalDocument;
};

export const visibleRangeForSource = (
  block: AddressableBlock,
  start: number,
  end: number,
): { start: number; end: number } | null => {
  if (!isTextLeaf(block) || end < start) return null;
  if (block.segments.some((segment) => segment.precision !== "exact"))
    return null;
  if (block.segments.length === 0) {
    const sourceStart = block.position.start.offset;
    if (
      block.sourceText !== visibleBlock(block) ||
      block.position.end.offset - sourceStart !== block.sourceText.length ||
      start < sourceStart ||
      end > block.position.end.offset
    )
      return null;
    return { start: start - sourceStart, end: end - sourceStart };
  }
  const sourceToVisible = (offset: number): number | null => {
    for (const segment of block.segments) {
      const width = segment.normalizedEnd - segment.normalizedStart;
      const sourceEnd = segment.sourceStartOffset + width;
      if (offset >= segment.sourceStartOffset && offset <= sourceEnd)
        return segment.normalizedStart + offset - segment.sourceStartOffset;
    }
    return null;
  };
  const visibleStart = sourceToVisible(start);
  const visibleEnd = sourceToVisible(end);
  if (visibleStart === null || visibleEnd === null || visibleEnd < visibleStart)
    return null;
  return { start: visibleStart, end: visibleEnd };
};

const sliceAttributionSpans = (
  spans: readonly AttributionSpan[],
  range: { start: number; end: number },
): readonly AttributionSpan[] =>
  spans
    .map((span) => {
      const start = Math.max(span.start, range.start);
      const end = Math.min(span.end, range.end);
      return end > start
        ? {
            start: start - range.start,
            end: end - range.start,
            attribution: span.attribution,
          }
        : null;
    })
    .filter((span): span is AttributionSpan => span !== null);
export const reattributeChangeSet = (
  changeSet: ChangeSet,
  provenance: ChangeSetProvenance,
): ChangeSet => {
  const baseDocumentBlocks = provenance.baseDocument
    ? new Map(
        flattenDocument(provenance.baseDocument).map((entry) => [
          entry.block.id,
          entry.block,
        ]),
      )
    : undefined;
  const headDocumentBlocks = provenance.headDocument
    ? new Map(
        flattenDocument(provenance.headDocument).map((entry) => [
          entry.block.id,
          entry.block,
        ]),
      )
    : undefined;
  const changes = changeSet.changes.map((change) => {
    let attribution: ChangeAttribution | undefined;
    let oldSpans: readonly AttributionSpan[] | undefined;
    let newSpans: readonly AttributionSpan[] | undefined;
    if ("blockId" in change) {
      oldSpans = provenance.baseBlocks.get(change.blockId);
      newSpans = provenance.headBlocks.get(change.blockId);
      attribution =
        provenance.headOperations.get(change.blockId) ??
        provenance.baseOperations.get(change.blockId) ??
        newSpans?.[0]?.attribution ??
        oldSpans?.[0]?.attribution;
    } else if (
      change.kind === "add-config" ||
      change.kind === "remove-config" ||
      change.kind === "replace-config"
    ) {
      attribution =
        (change.kind === "remove-config"
          ? (provenance.headConfigOperations?.get(change.path) ??
            provenance.baseConfigOperations?.get(change.path))
          : undefined) ??
        provenance.headConfig.get(change.path) ??
        provenance.baseConfig.get(change.path);
    } else if (
      change.kind === "add-dependency" ||
      change.kind === "remove-dependency" ||
      change.kind === "replace-dependency"
    ) {
      attribution =
        (change.kind === "remove-dependency"
          ? (provenance.headDependencyOperations?.get(change.key) ??
            provenance.baseDependencyOperations?.get(change.key))
          : undefined) ??
        provenance.headDependencies.get(change.key) ??
        provenance.baseDependencies.get(change.key);
    }
    const isTextChange =
      "blockId" in change &&
      (change.kind === "insert-text" ||
        change.kind === "delete-text" ||
        change.kind === "replace-text");
    if (isTextChange) {
      if (!baseDocumentBlocks?.has(change.blockId)) oldSpans = undefined;
      if (!headDocumentBlocks?.has(change.blockId)) newSpans = undefined;
    }
    if (!attribution && !oldSpans && !newSpans) return change;
    let updated: Change = attribution ? { ...change, attribution } : change;
    if (
      ("oldAttributionSpans" in updated || change.kind === "insert-text") &&
      oldSpans
    )
      updated = {
        ...updated,
        oldAttributionSpans: oldSpans,
      } as Change;
    if ("newAttributionSpans" in updated && newSpans)
      updated = { ...updated, newAttributionSpans: newSpans };
    return withChangeId(
      updated as unknown as Record<string, unknown>,
    ) as Change;
  });
  const withoutId = {
    schemaVersion: 1 as const,
    documentId: changeSet.documentId,
    baseRevision: changeSet.baseRevision,
    headRevision: changeSet.headRevision,
    changes,
    annotations: changeSet.annotations,
  };
  return { ...withoutId, id: `sha256:${canonicalHash(withoutId)}` };
};

const preserveDeltaAttribution = (
  changeSet: ChangeSet,
  base: LegalDocument,
  head: LegalDocument,
  baseBlocks: ReadonlyMap<string, readonly AttributionSpan[]>,
  attribution: ChangeAttribution,
): ChangeSet => {
  const baseById = new Map(
    flattenDocument(base).map((entry) => [entry.block.id, entry.block]),
  );
  const headById = new Map(
    flattenDocument(head).map((entry) => [entry.block.id, entry.block]),
  );
  const projectedTextSpans = new Map<string, readonly AttributionSpan[]>();
  for (const change of changeSet.changes) {
    if (
      change.kind !== "insert-text" &&
      change.kind !== "delete-text" &&
      change.kind !== "replace-text"
    )
      continue;
    const oldBlock = baseById.get(change.blockId);
    const newBlock = headById.get(change.blockId);
    if (!oldBlock || !newBlock) continue;
    const oldText = visibleBlock(oldBlock, base.metadata);
    const newText = visibleBlock(newBlock, head.metadata);
    const oldSpans =
      baseBlocks.get(change.blockId) ?? spansFor(oldText, attribution);
    projectedTextSpans.set(
      change.blockId,
      reattributeVisibleText(oldText, newText, oldSpans, attribution),
    );
  }
  const changes = changeSet.changes.map((change) => {
    if (!("blockId" in change)) return change;
    const oldBlock = baseById.get(change.blockId);
    const newBlock = headById.get(change.blockId);
    const oldText = oldBlock ? visibleBlock(oldBlock, base.metadata) : "";
    const newText = newBlock ? visibleBlock(newBlock, head.metadata) : "";
    const oldSpans =
      baseBlocks.get(change.blockId) ?? spansFor(oldText, attribution);
    let operationOldSpans = oldSpans;
    let operationNewSpans: readonly AttributionSpan[] | undefined;
    if (
      (change.kind === "insert-text" ||
        change.kind === "delete-text" ||
        change.kind === "replace-text") &&
      oldBlock &&
      newBlock
    ) {
      if ("oldSource" in change) {
        const range = visibleRangeForSource(
          oldBlock,
          change.oldSource.start,
          change.oldSource.end,
        );
        if (range) operationOldSpans = sliceAttributionSpans(oldSpans, range);
      }
      const projected =
        projectedTextSpans.get(change.blockId) ??
        reattributeVisibleText(oldText, newText, oldSpans, attribution);
      if ("newSource" in change) {
        const range = visibleRangeForSource(
          newBlock,
          change.newSource.start,
          change.newSource.end,
        );
        operationNewSpans = range
          ? sliceAttributionSpans(projected, range)
          : projected;
      } else {
        operationNewSpans = projected;
      }
    }
    let updated: Change = change;
    if ("oldAttributionSpans" in updated)
      updated = {
        ...updated,
        oldAttributionSpans: operationOldSpans,
      };
    if ("newAttributionSpans" in updated) {
      const newSpans =
        operationNewSpans ??
        (change.kind === "move-block"
          ? oldSpans
          : change.kind === "replace-container-shell"
            ? reattributeVisibleText(oldText, newText, oldSpans, attribution)
            : spansFor(newText, attribution));
      updated = { ...updated, newAttributionSpans: newSpans };
    }
    return withChangeId(
      updated as unknown as Record<string, unknown>,
    ) as Change;
  });
  const withoutId = {
    schemaVersion: 1 as const,
    documentId: changeSet.documentId,
    baseRevision: changeSet.baseRevision,
    headRevision: changeSet.headRevision,
    changes,
    annotations: changeSet.annotations,
  };
  return { ...withoutId, id: `sha256:${canonicalHash(withoutId)}` };
};

export const createRevisionDelta = (
  parent: {
    id: RevisionId;
    sourceObject: RevisionId;
    documentConfigObject: RevisionId;
    dependencyObjects: Readonly<Record<string, RevisionId>>;
  },
  documentId: string,
  base: LegalDocument,
  head: LegalDocument,
  baseAnnotations: readonly ReviewAnnotation[],
  headAnnotations: readonly ReviewAnnotation[],
  attribution: ChangeAttribution,
  baseConfig: JsonObject,
  headConfig: JsonObject,
  headDependencies: Readonly<Record<string, RevisionId>>,
  baseSource?: string,
  headSource?: string,
  baseBlocks?: ReadonlyMap<string, readonly AttributionSpan[]>,
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
    {
      baseConfig,
      headConfig,
      ...(parent.dependencyObjects !== undefined
        ? { baseDependencies: parent.dependencyObjects }
        : {}),
      ...(headDependencies !== undefined ? { headDependencies } : {}),
      ...(baseSource !== undefined ? { baseSource } : {}),
      ...(headSource !== undefined ? { headSource } : {}),
    },
  );
  const attributed = baseBlocks
    ? preserveDeltaAttribution(changeSet, base, head, baseBlocks, attribution)
    : changeSet;
  return {
    schemaVersion: 1,
    parentSourceObject: parent.sourceObject,
    parentDocumentConfigObject: parent.documentConfigObject,
    changes: attributed.changes,
    annotations: attributed.annotations,
  };
};

export const defaultAttribution = (
  actor: Actor,
  createdAt: string,
): ChangeAttribution => ({
  author: actor,
  createdAt,
});
export const rebaseOpenAnnotations = (
  base: LegalDocument,
  head: LegalDocument,
  annotations: readonly ReviewAnnotation[],
): readonly ReviewAnnotation[] => {
  const baseById = new Map(
    flattenDocument(base).map((entry) => [entry.block.id, entry.block]),
  );
  const headById = new Map(
    flattenDocument(head).map((entry) => [entry.block.id, entry.block]),
  );
  return annotations.map((annotation) => {
    if (annotation.status !== "open") return annotation;
    const previous = baseById.get(annotation.blockId);
    const next = headById.get(annotation.blockId);
    if (!previous || !next)
      throw new AgentDocxError(
        "ANNOTATION_CONFLICT",
        `Annotation block was deleted: ${annotation.id}`,
      );
    if (!annotation.range) return annotation;
    const oldText = visibleBlock(previous, base.metadata);
    const newText = visibleBlock(next, head.metadata);
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
      throw new AgentDocxError(
        "ANNOTATION_CONFLICT",
        `Annotation range is invalid: ${annotation.id}`,
      );
    const oldTokens = codePointTokens(oldText);
    const newTokens = codePointTokens(newText);
    const pairs = [...equalTokenPairs(oldTokens, newTokens).entries()].sort(
      ([, left], [, right]) => left - right,
    );
    const hunks: {
      oldStart: number;
      oldEnd: number;
      newStart: number;
      newEnd: number;
    }[] = [];
    let oldCursor = 0;
    let newCursor = 0;
    for (const [newIndex, oldIndex] of pairs) {
      if (oldIndex < oldCursor || newIndex < newCursor)
        throw new AgentDocxError(
          "ANNOTATION_CONFLICT",
          `Annotation edit script is invalid: ${annotation.id}`,
        );
      if (oldIndex > oldCursor || newIndex > newCursor) {
        const oldStart =
          oldCursor < oldTokens.length
            ? oldTokens[oldCursor]!.start
            : oldText.length;
        const newStart =
          newCursor < newTokens.length
            ? newTokens[newCursor]!.start
            : newText.length;
        hunks.push({
          oldStart,
          oldEnd:
            oldIndex > oldCursor ? oldTokens[oldIndex - 1]!.end : oldStart,
          newStart,
          newEnd:
            newIndex > newCursor ? newTokens[newIndex - 1]!.end : newStart,
        });
      }
      oldCursor = oldIndex + 1;
      newCursor = newIndex + 1;
    }
    if (oldCursor < oldTokens.length || newCursor < newTokens.length) {
      hunks.push({
        oldStart:
          oldCursor < oldTokens.length
            ? oldTokens[oldCursor]!.start
            : oldText.length,
        oldEnd: oldText.length,
        newStart:
          newCursor < newTokens.length
            ? newTokens[newCursor]!.start
            : newText.length,
        newEnd: newText.length,
      });
    }
    const translate = (offset: number, affinity: "right" | "left"): number => {
      let delta = 0;
      for (const hunk of hunks) {
        if (offset < hunk.oldStart) return offset + delta;
        if (offset > hunk.oldEnd) {
          delta += hunk.newEnd - hunk.newStart - (hunk.oldEnd - hunk.oldStart);
          continue;
        }
        if (hunk.oldStart === hunk.oldEnd && offset === hunk.oldStart)
          return affinity === "right" ? hunk.newEnd : hunk.newStart;
        if (offset === hunk.oldEnd) return hunk.newEnd;
        throw new AgentDocxError(
          "ANNOTATION_CONFLICT",
          `Annotation range overlaps an edit: ${annotation.id}`,
        );
      }
      return offset + delta;
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
