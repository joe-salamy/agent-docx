import type { Font } from "fontkit";
import LineBreaker from "linebreak";
import type { FlowBlock, InlineRun, NormalizedDocument } from "../markdown.js";
import type { LoadedFonts } from "../resolve.js";
import type {
  Diagnostic,
  LayoutProfile,
  ParagraphDiagnostic,
  TextStyle,
} from "../types.js";

export type PaginationOutput = {
  pageCount: number;
  equivalentPages: number;
  totalVisualLines: number;
  visualLinesByPage: number[];
  lastPage: {
    visualLines: number;
    usedTwips: number;
    usableTwips: number;
    bodyLineEquivalentsUsed: number;
    bodyLineCapacity: number;
  } | null;
  paragraphs: ParagraphDiagnostic[];
  warnings: Diagnostic[];
};

type WrappedLine = {
  used: number;
  available: number;
  height: number;
  text: string;
  block: FlowBlock;
  counted: boolean;
  start: number;
  end: number;
  footnoteRefs: string[];
};

type PlacedFootnoteLine = { footnoteId: string; line: WrappedLine };

type Page = {
  bodyUsed: number;
  footnoteUsed: number;
  visual: number;
  counted: number;
  bodyLines: WrappedLine[];
  footnoteLines: PlacedFootnoteLine[];
};

type PendingFootnote = {
  footnoteId: string;
  lines: readonly WrappedLine[];
  nextLine: number;
  warningEmitted: boolean;
};

type FootnoteState = {
  placed: Set<string>;
  pending: PendingFootnote[];
  relaxed: Set<string>;
};

type WrappedBlock = {
  block: FlowBlock;
  style: TextStyle;
  lines: WrappedLine[];
};

type Snapshot = {
  pages: Page[];
  page: Page;
  footnotes: FootnoteState;
};

const round = (number: number) =>
  Math.sign(number) * Math.floor(Math.abs(number) + 0.5);

const role = (run: InlineRun, style: TextStyle) =>
  (style.bold || run.bold
    ? style.italic || run.italic
      ? "boldItalic"
      : "bold"
    : style.italic || run.italic
      ? "italic"
      : "regular") as keyof Pick<
    LoadedFonts,
    "regular" | "bold" | "italic" | "boldItalic"
  >;

function width(font: Font, text: string, size: number) {
  const layout = font.layout(text);
  return round(
    (layout.positions.reduce((sum, position) => sum + position.xAdvance, 0) *
      size) /
      font.unitsPerEm,
  );
}

function candidateWidth(
  block: FlowBlock,
  start: number,
  end: number,
  fonts: LoadedFonts,
  style: TextStyle,
) {
  let cursor = 0;
  let total = 0;
  for (const run of block.runs) {
    const next = cursor + run.text.length;
    const from = Math.max(start, cursor);
    const to = Math.min(end, next);
    if (to > from) {
      total += width(
        fonts[role(run, style)].font,
        run.text.slice(from - cursor, to - cursor),
        style.fontSizeTwips,
      );
    }
    cursor = next;
  }
  return total;
}

function naturalHeight(
  fonts: LoadedFonts,
  block: FlowBlock,
  style: TextStyle,
  start: number,
  end: number,
) {
  let cursor = 0;
  let maximum = 0;
  for (const run of block.runs) {
    const next = cursor + run.text.length;
    if (Math.min(end, next) > Math.max(start, cursor)) {
      const font = fonts[role(run, style)].font;
      maximum = Math.max(
        maximum,
        round(
          ((font.ascent - font.descent + font.lineGap) * style.fontSizeTwips) /
            font.unitsPerEm,
        ),
      );
    }
    cursor = next;
  }
  return maximum || style.fontSizeTwips;
}

function linePitch(natural: number, style: TextStyle) {
  return style.lineSpacing.rule === "auto"
    ? round((natural * style.lineSpacing.numerator) / 240)
    : style.lineSpacing.rule === "exact"
      ? style.lineSpacing.twips
      : Math.max(natural, style.lineSpacing.twips);
}

function warnMissingGlyphs(
  block: FlowBlock,
  style: TextStyle,
  fonts: LoadedFonts,
  warnings: Diagnostic[],
) {
  for (const run of block.runs) {
    const font = fonts[role(run, style)].font;
    for (const character of run.text) {
      const codePoint = character.codePointAt(0)!;
      if (!font.hasGlyphForCodePoint(codePoint)) {
        warnings.push({
          code: "MISSING_GLYPH",
          severity: "warning",
          message: `The selected metric font has no glyph for U+${codePoint.toString(16).toUpperCase().padStart(4, "0")}.`,
          position: block.position,
        });
        return;
      }
    }
  }
}

function referencesInRange(block: FlowBlock, start: number, end: number) {
  const references: string[] = [];
  let cursor = 0;
  for (const run of block.runs) {
    const next = cursor + run.text.length;
    if (run.footnoteId !== undefined && cursor < end && next > start) {
      references.push(run.footnoteId);
    }
    cursor = next;
  }
  return references;
}

function wrap(
  block: FlowBlock,
  style: TextStyle,
  available: number,
  fonts: LoadedFonts,
  warnings: Diagnostic[],
  lineCapExclusions: LayoutProfile["pagination"]["lineCapExclusions"],
): WrappedLine[] {
  warnMissingGlyphs(block, style, fonts, warnings);
  const text = block.runs.map((run) => run.text).join("");
  const lines: WrappedLine[] = [];
  const counted =
    block.kind !== "footnote" && block.kind !== "blockquote"
      ? true
      : !lineCapExclusions.includes(block.kind);
  let segmentStart = 0;
  const mandatory = [...text.matchAll(/\n/g)].map(
    (match) => (match.index ?? 0) + 1,
  );
  mandatory.push(text.length);
  for (const segmentEndWithBreak of mandatory) {
    const segmentEnd =
      segmentEndWithBreak < text.length
        ? segmentEndWithBreak - 1
        : segmentEndWithBreak;
    let start = segmentStart;
    if (start === segmentEnd) {
      const end = Math.min(text.length, start + 1);
      const natural = naturalHeight(fonts, block, style, start, end);
      const lineAvailable =
        available -
        (lines.length === 0
          ? style.firstLineIndentTwips - style.hangingIndentTwips
          : 0);
      lines.push({
        used: 0,
        available: lineAvailable,
        height: linePitch(natural, style),
        text: "",
        block,
        counted,
        start,
        end: segmentEnd,
        footnoteRefs: referencesInRange(block, start, segmentEnd),
      });
    }
    while (start < segmentEnd) {
      const lineAvailable =
        available -
        (lines.length === 0
          ? style.firstLineIndentTwips - style.hangingIndentTwips
          : 0);
      const breaker = new LineBreaker(text.slice(start, segmentEnd));
      const opportunities: number[] = [];
      let nextBreak;
      while ((nextBreak = breaker.nextBreak()) !== null) {
        opportunities.push(start + nextBreak.position);
      }
      if (opportunities.at(-1) !== segmentEnd) opportunities.push(segmentEnd);
      const measuredEnd = (end: number) => {
        while (end > start && text[end - 1] === " ") end--;
        return end;
      };
      let chosen = start;
      for (const end of opportunities) {
        const used = candidateWidth(
          block,
          start,
          measuredEnd(end),
          fonts,
          style,
        );
        if (used <= lineAvailable) chosen = end;
        else break;
      }
      if (chosen === start) {
        chosen = opportunities[0] ?? segmentEnd;
        warnings.push({
          code: "UNBREAKABLE_OVERFLOW",
          severity: "warning",
          message: "An unbreakable token exceeds the available line width.",
          position: block.position,
        });
      }
      const contentEnd = measuredEnd(chosen);
      const used = candidateWidth(block, start, contentEnd, fonts, style);
      const natural = naturalHeight(fonts, block, style, start, contentEnd);
      lines.push({
        used,
        available: lineAvailable,
        height: linePitch(natural, style),
        text: text.slice(start, chosen).trimEnd(),
        block,
        counted,
        start,
        end: chosen,
        footnoteRefs: referencesInRange(block, start, chosen),
      });
      start = chosen;
      while (text[start] === " ") start++;
    }
    segmentStart = segmentEndWithBreak;
  }
  return lines;
}

function styleFor(profile: LayoutProfile, block: FlowBlock) {
  return block.kind === "heading"
    ? profile.headings[String(block.level ?? 1) as "1"]
    : block.kind === "blockquote"
      ? profile.blockquote
      : block.kind === "list"
        ? profile.list
        : block.kind === "footnote"
          ? profile.footnote
          : profile.body;
}

function emptyPage(): Page {
  return {
    bodyUsed: 0,
    footnoteUsed: 0,
    visual: 0,
    counted: 0,
    bodyLines: [],
    footnoteLines: [],
  };
}

function clonePage(page: Page): Page {
  return {
    ...page,
    bodyLines: [...page.bodyLines],
    footnoteLines: [...page.footnoteLines],
  };
}

function cloneFootnotes(state: FootnoteState): FootnoteState {
  return {
    placed: new Set(state.placed),
    pending: state.pending.map((pending) => ({ ...pending })),
    relaxed: new Set(state.relaxed),
  };
}

const occupied = (page: Page) => page.bodyUsed + page.footnoteUsed;
const hasContent = (page: Page) =>
  page.bodyLines.length > 0 || page.footnoteLines.length > 0;

export function paginate(
  document: NormalizedDocument,
  profile: LayoutProfile,
  fonts: LoadedFonts,
): PaginationOutput {
  if (document.blocks.length === 0) {
    return {
      pageCount: 0,
      equivalentPages: 0,
      totalVisualLines: 0,
      visualLinesByPage: [],
      lastPage: null,
      paragraphs: [],
      warnings: [],
    };
  }

  const usableWidth =
    profile.page.widthTwips -
    profile.page.marginsTwips.left -
    profile.page.marginsTwips.right -
    profile.page.gutterTwips;
  const usableHeight =
    profile.page.heightTwips -
    profile.page.marginsTwips.top -
    profile.page.marginsTwips.bottom;
  const warnings: Diagnostic[] = [];
  const bodyBlocks: WrappedBlock[] = document.blocks.map((block) => {
    if (block.kind === "pagebreak") {
      return { block, style: profile.body, lines: [] };
    }
    let style = styleFor(profile, block);
    const available =
      usableWidth - style.leftIndentTwips - style.rightIndentTwips;
    let lines = wrap(
      block,
      style,
      available,
      fonts,
      warnings,
      profile.pagination.lineCapExclusions,
    );
    if (
      profile.id === "frap-32" &&
      block.kind === "blockquote" &&
      lines.length >= 3
    ) {
      style = {
        ...style,
        lineSpacing: { rule: "auto", numerator: 240, denominator: 240 },
      };
      lines = wrap(
        block,
        style,
        available,
        fonts,
        warnings,
        profile.pagination.lineCapExclusions,
      );
    }
    return { block, style, lines };
  });
  const footnoteCache = new Map<string, readonly WrappedLine[]>();
  const wrappedFootnote = (id: string) => {
    let lines = footnoteCache.get(id);
    if (lines) return lines;
    const block = document.footnotes.get(id)!;
    const style = profile.footnote;
    lines = wrap(
      block,
      style,
      usableWidth - style.leftIndentTwips - style.rightIndentTwips,
      fonts,
      warnings,
      profile.pagination.lineCapExclusions,
    );
    footnoteCache.set(id, lines);
    return lines;
  };

  let pages: Page[] = [];
  let page = emptyPage();
  let footnotes: FootnoteState = {
    placed: new Set(),
    pending: [],
    relaxed: new Set(),
  };

  const snapshot = (): Snapshot => ({
    pages: pages.map(clonePage),
    page: clonePage(page),
    footnotes: cloneFootnotes(footnotes),
  });
  const restore = (saved: Snapshot) => {
    pages = saved.pages.map(clonePage);
    page = clonePage(saved.page);
    footnotes = cloneFootnotes(saved.footnotes);
  };
  const commitPage = () => {
    pages.push(page);
    page = emptyPage();
  };
  const lineFits = (target: Page, line: WrappedLine, extra = 0) => {
    const cap = profile.pagination.maxCountedLinesPerPage;
    return (
      occupied(target) + extra + line.height <= usableHeight &&
      (!line.counted || cap === null || target.counted < cap)
    );
  };
  const prefixThatFits = (
    target: Page,
    lines: readonly WrappedLine[],
    start: number,
    leadingSpacing: number,
  ) => {
    const probe = clonePage(target);
    let count = 0;
    for (let index = start; index < lines.length; index++) {
      const line = lines[index]!;
      const spacing = count === 0 ? leadingSpacing : 0;
      if (!lineFits(probe, line, spacing)) break;
      probe.footnoteUsed += spacing + line.height;
      probe.visual++;
      if (line.counted) probe.counted++;
      count++;
    }
    return count;
  };
  const footnoteFitsEmpty = (lines: readonly WrappedLine[]) =>
    prefixThatFits(emptyPage(), lines, 0, 0) === lines.length;
  const warnRelaxed = (
    state: FootnoteState,
    pending: PendingFootnote,
    diagnostics: Diagnostic[],
  ) => {
    if (state.relaxed.has(pending.footnoteId)) return;
    state.relaxed.add(pending.footnoteId);
    pending.warningEmitted = true;
    diagnostics.push({
      code: "FOOTNOTE_SPLIT_CONSTRAINT_RELAXED",
      severity: "warning",
      message:
        "Footnote constraints were relaxed to continue the definition across pages.",
      position: document.footnotes.get(pending.footnoteId)!.position,
    });
  };
  const enqueue = (state: FootnoteState, ids: readonly string[]) => {
    for (const id of ids) {
      if (state.placed.has(id)) continue;
      state.placed.add(id);
      state.pending.push({
        footnoteId: id,
        lines: wrappedFootnote(id),
        nextLine: 0,
        warningEmitted: state.relaxed.has(id),
      });
    }
  };
  const leadingFootnoteSpacing = (target: Page, pending: PendingFootnote) => {
    if (pending.nextLine !== 0) return 0;
    const previous = target.footnoteLines.at(-1);
    if (!previous || previous.footnoteId === pending.footnoteId) return 0;
    return Math.max(
      document.footnotes.get(previous.footnoteId)!
        ? profile.footnote.afterTwips
        : 0,
      profile.footnote.beforeTwips,
    );
  };
  const placePrefix = (
    target: Page,
    state: FootnoteState,
    pending: PendingFootnote,
    count: number,
    spacing: number,
  ) => {
    const discovered: string[] = [];
    for (let offset = 0; offset < count; offset++) {
      const line = pending.lines[pending.nextLine + offset]!;
      target.footnoteUsed += (offset === 0 ? spacing : 0) + line.height;
      target.visual++;
      if (line.counted) target.counted++;
      target.footnoteLines.push({ footnoteId: pending.footnoteId, line });
      discovered.push(...line.footnoteRefs);
    }
    pending.nextLine += count;
    if (pending.nextLine === pending.lines.length) state.pending.shift();
    enqueue(state, discovered);
  };
  const placePendingHead = (
    target: Page,
    state: FootnoteState,
    diagnostics: Diagnostic[],
  ) => {
    const pending = state.pending[0];
    if (!pending) return 0;
    const spacing = leadingFootnoteSpacing(target, pending);
    const remaining = pending.lines.length - pending.nextLine;
    const maximum = prefixThatFits(
      target,
      pending.lines,
      pending.nextLine,
      spacing,
    );
    const wholeFitsEmpty = footnoteFitsEmpty(pending.lines);
    if (!wholeFitsEmpty) warnRelaxed(state, pending, diagnostics);
    if (maximum >= remaining) {
      placePrefix(target, state, pending, remaining, spacing);
      return remaining;
    }
    if (
      pending.nextLine === 0 &&
      profile.footnote.keepLines &&
      wholeFitsEmpty
    ) {
      return 0;
    }

    const orphan = profile.pagination.orphanLines;
    const widow = profile.pagination.widowLines;
    let constrained = 0;
    for (let count = maximum; count > 0; count--) {
      if (count >= orphan && remaining - count >= widow) {
        constrained = count;
        break;
      }
    }
    if (constrained > 0) {
      placePrefix(target, state, pending, constrained, spacing);
      return constrained;
    }

    const emptyMaximum = prefixThatFits(
      emptyPage(),
      pending.lines,
      pending.nextLine,
      0,
    );
    let emptyConstrained = false;
    for (let count = emptyMaximum; count > 0; count--) {
      if (count >= orphan && remaining - count >= widow) {
        emptyConstrained = true;
        break;
      }
    }
    if (emptyConstrained || maximum === 0) return 0;

    warnRelaxed(state, pending, diagnostics);
    placePrefix(target, state, pending, maximum, spacing);
    return maximum;
  };
  const reserveOnCurrentPage = (
    target: Page,
    state: FootnoteState,
    diagnostics: Diagnostic[],
  ) => {
    let placed = 0;
    while (state.pending.length) {
      const head = state.pending[0]!;
      const count = placePendingHead(target, state, diagnostics);
      placed += count;
      if (count === 0 || state.pending[0] === head) break;
    }
    return placed;
  };
  const forceOneFootnoteLine = (
    target: Page,
    state: FootnoteState,
    diagnostics: Diagnostic[],
  ) => {
    const pending = state.pending[0]!;
    warnRelaxed(state, pending, diagnostics);
    placePrefix(target, state, pending, 1, 0);
  };
  const drainContinuations = () => {
    if (!footnotes.pending.length) return;
    commitPage();
    while (footnotes.pending.length) {
      const diagnostics: Diagnostic[] = [];
      const count = reserveOnCurrentPage(page, footnotes, diagnostics);
      if (count === 0) forceOneFootnoteLine(page, footnotes, diagnostics);
      warnings.push(...diagnostics);
      if (footnotes.pending.length) commitPage();
    }
  };
  const attemptBodyLine = (
    line: WrappedLine,
    spacing: number,
  ): {
    page: Page;
    footnotes: FootnoteState;
    diagnostics: Diagnostic[];
  } | null => {
    const target = clonePage(page);
    const state = cloneFootnotes(footnotes);
    const diagnostics: Diagnostic[] = [];
    const occupiedBefore = occupied(target);
    if (!lineFits(target, line, spacing)) {
      if (hasContent(target)) return null;
      target.bodyUsed += spacing + line.height;
    } else {
      target.bodyUsed += spacing + line.height;
    }
    target.visual++;
    if (line.counted) target.counted++;
    target.bodyLines.push(line);
    const pendingBefore = state.pending.length;
    enqueue(state, line.footnoteRefs);
    const footnoteLinesBefore = target.footnoteLines.length;
    reserveOnCurrentPage(target, state, diagnostics);
    if (
      state.pending.length > pendingBefore &&
      target.footnoteLines.length === footnoteLinesBefore &&
      occupiedBefore > 0
    ) {
      return null;
    }
    return { page: target, footnotes: state, diagnostics };
  };
  const simulateUnit = (
    unit: readonly { line: WrappedLine; spacing: number }[],
    startPage: Page,
    startFootnotes: FootnoteState,
  ) => {
    const target = clonePage(startPage);
    const state = cloneFootnotes(startFootnotes);
    const diagnostics: Diagnostic[] = [];
    for (const { line, spacing } of unit) {
      if (!lineFits(target, line, spacing)) return false;
      target.bodyUsed += spacing + line.height;
      target.visual++;
      if (line.counted) target.counted++;
      target.bodyLines.push(line);
      enqueue(state, line.footnoteRefs);
      reserveOnCurrentPage(target, state, diagnostics);
    }
    return state.pending.length === 0;
  };
  const unitFromBlocks = (
    indexes: readonly number[],
    terminatingLines: number | null,
    firstPriorAfter: number,
    startHasBody: boolean,
  ) => {
    const unit: { line: WrappedLine; spacing: number }[] = [];
    let previousAfter = firstPriorAfter;
    let hasBody = startHasBody;
    for (let offset = 0; offset < indexes.length; offset++) {
      const record = bodyBlocks[indexes[offset]!]!;
      const lineCount =
        offset === indexes.length - 1 && terminatingLines !== null
          ? Math.min(terminatingLines, record.lines.length)
          : record.lines.length;
      for (let lineIndex = 0; lineIndex < lineCount; lineIndex++) {
        unit.push({
          line: record.lines[lineIndex]!,
          spacing:
            lineIndex === 0 && hasBody
              ? Math.max(previousAfter, record.style.beforeTwips)
              : 0,
        });
        hasBody = true;
      }
      previousAfter = record.style.afterTwips;
    }
    return unit;
  };

  let priorAfter = 0;
  for (let blockIndex = 0; blockIndex < bodyBlocks.length; blockIndex++) {
    const record = bodyBlocks[blockIndex]!;
    if (record.block.kind === "pagebreak") {
      commitPage();
      priorAfter = 0;
      continue;
    }

    if (record.style.keepWithNext) {
      const indexes = [blockIndex];
      let cursor = blockIndex;
      while (
        bodyBlocks[cursor]!.style.keepWithNext &&
        cursor + 1 < bodyBlocks.length &&
        bodyBlocks[cursor + 1]!.block.kind !== "pagebreak"
      ) {
        cursor++;
        indexes.push(cursor);
      }
      const terminator = bodyBlocks[indexes.at(-1)!]!;
      const hasTerminatingBlock = !terminator.style.keepWithNext;
      if (hasTerminatingBlock && indexes.length > 1) {
        const terminatingLines = terminator.style.keepLines
          ? terminator.lines.length
          : 1;
        const currentUnit = unitFromBlocks(
          indexes,
          terminatingLines,
          priorAfter,
          page.bodyLines.length > 0,
        );
        const emptyUnit = unitFromBlocks(indexes, terminatingLines, 0, false);
        if (
          simulateUnit(emptyUnit, emptyPage(), footnotes) &&
          !simulateUnit(currentUnit, page, footnotes) &&
          hasContent(page)
        ) {
          commitPage();
        }
      }
    }

    if (record.style.keepLines) {
      const currentUnit = unitFromBlocks(
        [blockIndex],
        null,
        priorAfter,
        page.bodyLines.length > 0,
      );
      const emptyUnit = unitFromBlocks([blockIndex], null, 0, false);
      if (
        simulateUnit(emptyUnit, emptyPage(), footnotes) &&
        !simulateUnit(currentUnit, page, footnotes) &&
        hasContent(page)
      ) {
        commitPage();
      }
    }

    if (
      !record.style.keepLines &&
      record.lines.length > 1 &&
      hasContent(page)
    ) {
      const firstLine = unitFromBlocks(
        [blockIndex],
        1,
        priorAfter,
        page.bodyLines.length > 0,
      );
      const orphanLines = unitFromBlocks(
        [blockIndex],
        Math.min(profile.pagination.orphanLines, record.lines.length),
        priorAfter,
        page.bodyLines.length > 0,
      );
      if (
        simulateUnit(firstLine, page, footnotes) &&
        !simulateUnit(orphanLines, page, footnotes)
      ) {
        commitPage();
      }
    }

    const beforeLine: Snapshot[] = [];
    let lineIndex = 0;
    while (lineIndex < record.lines.length) {
      beforeLine[lineIndex] = snapshot();
      const spacing =
        lineIndex === 0 && page.bodyLines.length > 0
          ? Math.max(priorAfter, record.style.beforeTwips)
          : 0;
      const result = attemptBodyLine(record.lines[lineIndex]!, spacing);
      if (!result) {
        const remaining = record.lines.length - lineIndex;
        const placedOnPage = page.bodyLines.filter(
          (line) => line.block === record.block,
        ).length;
        if (
          remaining < profile.pagination.widowLines &&
          placedOnPage >= profile.pagination.orphanLines
        ) {
          const move = Math.min(profile.pagination.orphanLines, placedOnPage);
          const targetIndex = lineIndex - move;
          restore(beforeLine[targetIndex]!);
          commitPage();
          lineIndex = targetIndex;
          continue;
        }
        commitPage();
        continue;
      }
      page = result.page;
      footnotes = result.footnotes;
      warnings.push(...result.diagnostics);
      lineIndex++;
      drainContinuations();
    }
    priorAfter = record.style.afterTwips;
  }

  if (hasContent(page) || pages.length === 0) pages.push(page);
  const paragraphResults: ParagraphDiagnostic[] = [];
  let paragraphIndex = 0;
  for (const block of document.blocks) {
    if (block.kind === "pagebreak") continue;
    const occurrences: { page: number; line: WrappedLine }[] = [];
    pages.forEach((placedPage, pageIndex) =>
      placedPage.bodyLines.forEach((line) => {
        if (line.block === block) {
          occurrences.push({ page: pageIndex + 1, line });
        }
      }),
    );
    if (occurrences.length) {
      const last = occurrences.at(-1)!;
      paragraphResults.push({
        source: "deterministic",
        index: paragraphIndex++,
        position: block.position,
        startPage: occurrences[0]!.page,
        endPage: last.page,
        visualLines: occurrences.length,
        lastLineUsedTwips: last.line.used,
        lastLineAvailableTwips: last.line.available,
        lastLineRatio:
          last.line.available === 0 ? 0 : last.line.used / last.line.available,
        preview: block.runs
          .map((run) => run.text)
          .join("")
          .replace(/\s+/g, " ")
          .trim()
          .slice(0, 80),
      });
    }
  }

  const last = pages.at(-1)!;
  const representative = document.blocks.find(
    (block) => block.kind !== "pagebreak",
  );
  const bodyNatural = naturalHeight(
    fonts,
    {
      ...(representative ?? document.blocks[0]!),
      runs: [{ text: "Ag", bold: false, italic: false }],
    },
    profile.body,
    0,
    2,
  );
  const bodyPitch = linePitch(bodyNatural, profile.body);
  return {
    pageCount: pages.length,
    equivalentPages: pages.length - 1 + occupied(last) / usableHeight,
    totalVisualLines: pages.reduce(
      (total, placedPage) => total + placedPage.visual,
      0,
    ),
    visualLinesByPage: pages.map((placedPage) => placedPage.visual),
    lastPage: {
      visualLines: last.visual,
      usedTwips: occupied(last),
      usableTwips: usableHeight,
      bodyLineEquivalentsUsed: occupied(last) / bodyPitch,
      bodyLineCapacity: Math.floor(usableHeight / bodyPitch),
    },
    paragraphs: paragraphResults,
    warnings,
  };
}
