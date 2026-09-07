import { createHash } from "node:crypto";
import canonicalize from "canonicalize";
import { AgentDocxError, type JsonValue } from "../types.js";
import type {
  AddressableBlock,
  BlockId,
  LegalBlock,
  LegalDocument,
  RevisionId,
} from "../legal/model.js";
import type {
  AttributionSpan,
  BlockLocation,
  Change,
  ChangeAttribution,
  ContainerShell,
} from "./types.js";
import type { ChangeSetOptions, JsonObject } from "./diff.js";

export const compareText = (left: string, right: string): number =>
  left < right ? -1 : left > right ? 1 : 0;

export const jsonObject = (value: unknown): JsonObject => value as JsonObject;

export const pointerSegment = (value: string): string =>
  value.replaceAll("~", "~0").replaceAll("/", "~1");

export type PresentValue =
  | { present: false }
  | { present: true; value: JsonValue };

export const configChanges = (
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

export const dependencyChanges = (
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

export const configAndDependencyChanges = (
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

export type FlatBlock = {
  block: AddressableBlock;
  location: BlockLocation;
};

export const canonicalHash = (value: unknown): string => {
  const serialized = canonicalize(value);
  if (serialized === undefined)
    throw new AgentDocxError("INTERNAL_ERROR", "Cannot canonicalize change");
  return createHash("sha256").update(serialized).digest("hex");
};

export const withoutPositions = (value: unknown): unknown => {
  if (Array.isArray(value)) return value.map(withoutPositions);
  if (value !== null && typeof value === "object")
    return Object.fromEntries(
      Object.entries(value as Record<string, unknown>)
        .filter(([key]) => key !== "position" && key !== "sourceStartOffset")
        .map(([key, entry]) => [key, withoutPositions(entry)]),
    );
  return value;
};

export const canonicalBlock = (block: AddressableBlock): string => {
  const serialized = canonicalize(withoutPositions(block));
  if (serialized === undefined)
    throw new AgentDocxError("INTERNAL_ERROR", "Cannot canonicalize block");
  return serialized;
};

export const withChangeId = <const Value extends Record<string, unknown>>(
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

export const spansFor = (
  text: string,
  attribution: ChangeAttribution,
): readonly AttributionSpan[] =>
  text.length === 0 ? [] : [{ start: 0, end: text.length, attribution }];
export type TextToken = { start: number; end: number; text: string };

export const tokenizeVisibleText = (text: string): readonly TextToken[] => {
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
export const codePointTokens = (text: string): readonly TextToken[] => {
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
export const MAX_DIFF_TOKENS = 50_000;
export const MAX_DIFF_TRACE_CELLS = 4_000_000;

export const enforceTokenBudget = (
  oldTokens: readonly TextToken[],
  newTokens: readonly TextToken[],
): void => {
  if (oldTokens.length + newTokens.length > MAX_DIFF_TOKENS)
    throw new AgentDocxError(
      "DIFF_TOO_LARGE",
      `Revision diff exceeds the ${MAX_DIFF_TOKENS}-token budget`,
    );
};

export const equalTokenPairs = (
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

export const sameAttribution = (
  left: ChangeAttribution,
  right: ChangeAttribution,
): boolean => canonicalize(left) === canonicalize(right);

export const appendAttributionSpan = (
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

export const flatten = (
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

export const flattenDocument = (document: LegalDocument): FlatBlock[] => {
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

export const blockSourceRange = (block: AddressableBlock) => ({
  start: block.position.start.offset,
  end: block.position.end.offset,
  text: block.sourceText,
});

export const structuralSourceRange = (
  block: AddressableBlock,
  source?: string,
) => {
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
export const shellSourceRanges = (
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
export const shellFor = (
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
export const containerShellSignature = (
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

export const isContainer = (
  block: AddressableBlock,
): block is Extract<
  LegalBlock,
  { kind: "list" | "exhibit" | "length-exclusion" }
> =>
  block.kind === "list" ||
  block.kind === "exhibit" ||
  block.kind === "length-exclusion";
