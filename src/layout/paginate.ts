import LineBreaker from "linebreak";
import type {
  FlowBlock,
  NormalizedDocument,
  SectionIndex,
  TableFlowBlock,
  TextFlowBlock,
} from "../markdown.js";
import type { LoadedFonts } from "../resolve.js";
import { flowStyleFor } from "./style.js";
import { candidateWidth, role, round, tableColumnWidths } from "./table.js";
import {
  applyWidowOrphan,
  buildDiagnostics,
  placeKeepUnit,
  wrapBodyBlocks,
  wrapFootnotes,
  type FootnoteState,
  type Page,
  type PendingFootnote,
  type Snapshot,
  type WrappedBlock,
  type WrappedLine,
  type WrappedTable,
} from "./wrap.js";
import type { Diagnostic, SourcePosition } from "../types.js";
import type { LayoutProfile, TextStyle } from "./profile.js";
import type { ParagraphDiagnostic, SectionDiagnostic } from "../measurement.js";

export type PaginationOutput = {
  pageCount: number;
  equivalentPages: number;
  totalVisualLines: number;
  visualLinesByPage: number[];
  countedLinesByPage: number[];
  lastPage: {
    visualLines: number;
    usedTwips: number;
    usableTwips: number;
    bodyLineEquivalentsUsed: number;
    bodyLineCapacity: number;
  } | null;
  paragraphs: ParagraphDiagnostic[];
  warnings: Diagnostic[];
  sections?: SectionDiagnostic[];
};

function naturalHeight(
  fonts: LoadedFonts,
  block: TextFlowBlock,
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
  block: TextFlowBlock,
  style: TextStyle,
  fonts: LoadedFonts,
  warnings: Diagnostic[],
) {
  for (const run of block.runs) {
    const font = fonts[role(run, style)].font;
    for (const character of run.text) {
      if (character === "\t") continue;
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

function referencesInRange(block: TextFlowBlock, start: number, end: number) {
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
const advanceSourcePoint = (
  point: SourcePosition["start"],
  text: string,
): SourcePosition["start"] => {
  let line = point.line;
  let column = point.column;
  for (const character of text) {
    if (character === "\n") {
      line++;
      column = 1;
    } else {
      column++;
    }
  }
  return { line, column, offset: point.offset + text.length };
};

const sourceRangesFor = (
  block: TextFlowBlock,
  start: number,
  end: number,
): readonly {
  position: SourcePosition;
  precision: "exact" | "node";
}[] => {
  const selected =
    start === end
      ? block.sourceSegments.filter(
          (segment) =>
            segment.normalizedEnd === start ||
            (segment.normalizedStart <= start && segment.normalizedEnd > start),
        )
      : block.sourceSegments.filter(
          (segment) =>
            segment.normalizedStart < end && segment.normalizedEnd > start,
        );
  return selected.map((segment) => {
    if (segment.precision === "node")
      return { position: segment.position, precision: segment.precision };
    const from = Math.max(start, segment.normalizedStart);
    const to = Math.min(end, segment.normalizedEnd);
    const before = block.normalizedText.slice(segment.normalizedStart, from);
    const selectedText = block.normalizedText.slice(from, to);
    const rangeStart = advanceSourcePoint(segment.position.start, before);
    return {
      position: {
        start: rangeStart,
        end: advanceSourcePoint(rangeStart, selectedText),
      },
      precision: segment.precision,
    };
  });
};

function wrap(
  block: TextFlowBlock,
  style: TextStyle,
  available: number,
  fonts: LoadedFonts,
  warnings: Diagnostic[],
  lineCapExclusions: LayoutProfile["pagination"]["lineCapExclusions"],
): WrappedLine[] {
  warnMissingGlyphs(block, style, fonts, warnings);
  const text = block.normalizedText;
  const lines: WrappedLine[] = [];
  const counted =
    block.kind !== "footnote" && block.kind !== "blockquote"
      ? true
      : !lineCapExclusions.includes(block.kind);
  const hardBreaks = [...text.matchAll(/\n/g)].map((match) => match.index ?? 0);
  const segmentEnds = [...hardBreaks, text.length];
  let segmentStart = 0;
  for (
    let segmentIndex = 0;
    segmentIndex < segmentEnds.length;
    segmentIndex++
  ) {
    const segmentEnd = segmentEnds[segmentIndex]!;
    let start = segmentStart;
    const segmentCause: WrappedLine["startCause"] =
      segmentIndex === 0 ? "start" : "hard";
    if (start === segmentEnd) {
      const natural = naturalHeight(
        fonts,
        block,
        style,
        Math.max(0, start - 1),
        start,
      );
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
        visualCount: 1,
        countedCount: counted ? 1 : 0,
        start,
        end: segmentEnd,
        contentEnd: segmentEnd,
        startCause: segmentCause,
        overflowed: false,
        footnoteRefs: [],
      });
    }
    let lineInSegment = 0;
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
      const measuredEnd = (candidateEnd: number) => {
        while (candidateEnd > start && text[candidateEnd - 1] === " ")
          candidateEnd--;
        return candidateEnd;
      };
      let chosen = start;
      for (const candidateEnd of opportunities) {
        const used = candidateWidth(
          block,
          start,
          measuredEnd(candidateEnd),
          fonts,
          style,
        );
        if (used <= lineAvailable) chosen = candidateEnd;
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
        text: text.slice(start, contentEnd),
        block,
        counted,
        visualCount: 1,
        countedCount: counted ? 1 : 0,
        start,
        end: chosen,
        contentEnd,
        startCause: lineInSegment === 0 ? segmentCause : "soft",
        overflowed: used > lineAvailable,
        footnoteRefs: referencesInRange(block, start, chosen),
      });
      lineInSegment++;
      start = chosen;
      while (text[start] === " ") start++;
    }
    segmentStart = segmentEnd + 1;
  }
  return lines;
}

function wrapTable(
  block: TableFlowBlock,
  profile: LayoutProfile,
  fonts: LoadedFonts,
  usableWidth: number,
  usableHeight: number,
  warnings: Diagnostic[],
): WrappedTable {
  const gridWidths = tableColumnWidths(block, profile, fonts, usableWidth);
  const lines: WrappedLine[] = [];
  let headerLineCount = 0;
  let relaxed = false;
  for (let rowIndex = 0; rowIndex < block.rows.length; rowIndex++) {
    const row = block.rows[rowIndex]!;
    const style = rowIndex === 0 ? profile.table.header : profile.table.body;
    const cells = row.map((cell, columnIndex) => {
      const cellBlock: TextFlowBlock = {
        kind: "paragraph",
        runs: cell.runs,
        normalizedText: cell.normalizedText,
        sourceSegments: cell.sourceSegments,
        position: cell.position,
        footnoteRefs: cell.footnoteRefs,
      };
      const available =
        gridWidths[columnIndex]! -
        profile.table.cellPaddingTwips.left -
        profile.table.cellPaddingTwips.right -
        2 * profile.table.borderTwips -
        style.leftIndentTwips -
        style.rightIndentTwips;
      return wrap(
        cellBlock,
        style,
        available,
        fonts,
        warnings,
        profile.pagination.lineCapExclusions,
      );
    });
    const bandCount = Math.max(
      1,
      ...cells.map((cellLines) => cellLines.length),
    );
    const rowLines: WrappedLine[] = [];
    for (let bandIndex = 0; bandIndex < bandCount; bandIndex++) {
      const top =
        bandIndex === 0
          ? profile.table.cellPaddingTwips.top +
            profile.table.borderTwips +
            style.beforeTwips
          : 0;
      const bottom =
        bandIndex === bandCount - 1
          ? profile.table.cellPaddingTwips.bottom +
            profile.table.borderTwips +
            style.afterTwips
          : 0;
      const bandHeight = Math.max(
        0,
        ...cells.map((cellLines) => cellLines[bandIndex]?.height ?? 0),
      );
      const bandFootnoteRefs = [
        ...new Set(
          cells.flatMap(
            (cellLines) => cellLines[bandIndex]?.footnoteRefs ?? [],
          ),
        ),
      ];
      rowLines.push({
        used: usableWidth,
        available: usableWidth,
        height: top + bandHeight + bottom,
        text: "",
        block,
        counted: true,
        visualCount: 1,
        countedCount: 1,
        start: 0,
        end: 0,
        contentEnd: 0,
        startCause: "start",
        overflowed: false,
        footnoteRefs: bandFootnoteRefs,
        rowStart: bandIndex === 0,
        rowEnd: bandIndex === bandCount - 1,
      });
    }
    if (
      rowLines.reduce((total, line) => total + line.height, 0) > usableHeight &&
      !relaxed
    ) {
      relaxed = true;
      warnings.push({
        code: "TABLE_ROW_SPLIT_CONSTRAINT_RELAXED",
        severity: "warning",
        message:
          "An intrinsically oversized table row was split across page bands.",
        position: block.position,
      });
    }
    lines.push(...rowLines);
    if (rowIndex === 0) headerLineCount = lines.length;
  }
  return { lines, headerLineCount, gridWidths };
}

function styleFor(profile: LayoutProfile, block: TextFlowBlock) {
  return flowStyleFor(block, profile).style;
}

function emptyPage(trackSections = false): Page {
  return {
    bodyUsed: 0,
    footnoteUsed: 0,
    visual: 0,
    counted: 0,
    bodyLines: [],
    footnoteLines: [],
    ...(trackSections ? { sectionTouches: new Set<number>() } : {}),
  };
}

function clonePage(page: Page): Page {
  return {
    ...page,
    bodyLines: [...page.bodyLines],
    footnoteLines: [...page.footnoteLines],
    ...(page.sectionTouches
      ? { sectionTouches: new Set(page.sectionTouches) }
      : {}),
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
  sectionIndex?: SectionIndex,
  usableHeightOverride?: number,
): PaginationOutput {
  if (document.blocks.length === 0) {
    return {
      pageCount: 0,
      equivalentPages: 0,
      totalVisualLines: 0,
      visualLinesByPage: [],
      countedLinesByPage: [],
      lastPage: null,
      paragraphs: [],
      warnings: [],
      ...(sectionIndex
        ? {
            sections: sectionIndex.sections.map((section) => ({
              source: "deterministic" as const,
              index: section.index,
              parentIndex: section.parentIndex,
              heading: section.heading,
              position: section.position,
              empty: section.empty,
              startPage: null,
              endPage: null,
              pageCount: 0,
              bodyVisualLines: 0,
              footnoteVisualLines: 0,
              visualLines: 0,
              countedLines: 0,
              pages: [],
            })),
          }
        : {}),
    };
  }

  const usableWidth =
    profile.page.widthTwips -
    profile.page.marginsTwips.left -
    profile.page.marginsTwips.right -
    profile.page.gutterTwips;
  const usableHeight =
    usableHeightOverride ??
    profile.page.heightTwips -
      profile.page.marginsTwips.top -
      profile.page.marginsTwips.bottom;
  const warnings: Diagnostic[] = [];
  const headerRepeatWarnings = new Set<FlowBlock>();
  const { bodyBlocks, bodyPitch } = wrapBodyBlocks(
    document,
    profile,
    fonts,
    usableWidth,
    usableHeight,
    warnings,
    wrap,
    wrapTable,
    styleFor,
    naturalHeight,
    linePitch,
  );
  const wrappedFootnote = wrapFootnotes(
    document,
    profile,
    usableWidth,
    fonts,
    warnings,
    wrap,
  );

  let pages: Page[] = [];
  let page = emptyPage(sectionIndex !== undefined);
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
  const restore = (saved: Snapshot) => ({
    pages: saved.pages.map(clonePage),
    page: clonePage(saved.page),
    footnotes: cloneFootnotes(saved.footnotes),
  });
  const commitRestored = (saved: {
    pages: Page[];
    page: Page;
    footnotes: FootnoteState;
  }) => {
    saved.pages.push(saved.page);
    return { ...saved, page: emptyPage(sectionIndex !== undefined) };
  };
  const commitTargetPage = (target: Page) => {
    pages.push(target);
    return emptyPage(sectionIndex !== undefined);
  };
  const commitPage = () => {
    pages.push(page);
    page = emptyPage(sectionIndex !== undefined);
  };
  const lineFits = (target: Page, line: WrappedLine, extra = 0) => {
    const cap = profile.pagination.maxCountedLinesPerPage;
    return (
      occupied(target) + extra + line.height <= usableHeight &&
      (cap === null || target.counted + line.countedCount <= cap)
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
      probe.visual += line.visualCount;
      probe.counted += line.countedCount;
      count++;
    }
    return count;
  };
  const footnoteFitsEmpty = (lines: readonly WrappedLine[]) =>
    prefixThatFits(emptyPage(sectionIndex !== undefined), lines, 0, 0) ===
    lines.length;
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
  const enqueue = (
    state: FootnoteState,
    ids: readonly string[],
    ownerIndex: number,
  ) => {
    for (const id of ids) {
      if (state.placed.has(id)) continue;
      state.placed.add(id);
      state.pending.push({
        footnoteId: id,
        blocks: wrappedFootnote(id),
        blockIndex: 0,
        nextLine: 0,
        warningEmitted: state.relaxed.has(id),
        ownerIndex,
      });
    }
  };
  const leadingFootnoteSpacing = (target: Page, pending: PendingFootnote) => {
    if (pending.nextLine !== 0) return 0;
    if (pending.blockIndex > 0)
      return Math.max(
        profile.footnote.afterTwips,
        profile.footnote.beforeTwips,
      );
    const previous = target.footnoteLines.at(-1);
    if (!previous || previous.footnoteId === pending.footnoteId) return 0;
    return Math.max(profile.footnote.afterTwips, profile.footnote.beforeTwips);
  };
  const placePrefix = (
    target: Page,
    state: FootnoteState,
    pending: PendingFootnote,
    count: number,
    spacing: number,
  ) => {
    const discovered: string[] = [];
    const current = pending.blocks[pending.blockIndex]!;
    const ownerIndex = pending.ownerIndex;
    for (let offset = 0; offset < count; offset++) {
      const line = current[pending.nextLine + offset]!;
      target.footnoteUsed += (offset === 0 ? spacing : 0) + line.height;
      target.visual += line.visualCount;
      target.counted += line.countedCount;
      target.footnoteLines.push({
        footnoteId: pending.footnoteId,
        line,
        ownerIndex,
      });
      discovered.push(...line.footnoteRefs);
    }
    pending.nextLine += count;
    if (pending.nextLine === current.length) {
      pending.blockIndex++;
      pending.nextLine = 0;
      if (pending.blockIndex === pending.blocks.length) state.pending.shift();
    }
    enqueue(state, discovered, ownerIndex);
  };
  const placePendingHead = (
    target: Page,
    state: FootnoteState,
    diagnostics: Diagnostic[],
  ) => {
    const pending = state.pending[0];
    if (!pending) return 0;
    const current = pending.blocks[pending.blockIndex]!;
    const spacing = leadingFootnoteSpacing(target, pending);
    const remaining = current.length - pending.nextLine;
    const maximum = prefixThatFits(target, current, pending.nextLine, spacing);
    const wholeFitsEmpty = footnoteFitsEmpty(current);
    if (!wholeFitsEmpty) warnRelaxed(state, pending, diagnostics);
    if (maximum >= remaining) {
      const next = pending.blocks[pending.blockIndex + 1];
      if (
        pending.nextLine === 0 &&
        profile.footnote.keepWithNext &&
        next &&
        prefixThatFits(target, [...current, ...next.slice(0, 1)], 0, spacing) <=
          current.length
      ) {
        return 0;
      }
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

    if (!profile.pagination.widowOrphanControl) {
      if (maximum === 0) return 0;
      placePrefix(target, state, pending, maximum, spacing);
      return maximum;
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
      emptyPage(sectionIndex !== undefined),
      current,
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
      const blockIndex = head.blockIndex;
      const count = placePendingHead(target, state, diagnostics);
      placed += count;
      if (count === 0) break;
      if (
        state.pending[0] === head &&
        head.blockIndex === blockIndex &&
        head.nextLine > 0
      )
        break;
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
    target.visual += line.visualCount;
    target.counted += line.countedCount;
    target.bodyLines.push(line);
    enqueue(
      state,
      line.footnoteRefs,
      sectionIndex?.deepestOwnerByBlock.get(line.block) ?? 0,
    );
    const footnoteLinesBefore = target.footnoteLines.length;
    reserveOnCurrentPage(target, state, diagnostics);
    if (
      state.pending.length > 0 &&
      target.footnoteLines.length === footnoteLinesBefore &&
      occupiedBefore > 0
    ) {
      return null;
    }
    return { page: target, footnotes: state, diagnostics };
  };
  const warnTableHeaderRelaxed = (record: WrappedBlock) => {
    if (headerRepeatWarnings.has(record.block)) return;
    headerRepeatWarnings.add(record.block);
    warnings.push({
      code: "TABLE_HEADER_REPEAT_CONSTRAINT_RELAXED",
      severity: "warning",
      message:
        "The table header could not repeat with a following content band.",
      position: record.block.position,
    });
  };
  const repeatTableHeader = (record: WrappedBlock, lineIndex: number) => {
    if (
      record.block.kind !== "table" ||
      !profile.table.repeatHeader ||
      !record.headerLineCount ||
      lineIndex < record.headerLineCount
    )
      return;
    const headers = record.lines.slice(0, record.headerLineCount);
    let rowEnd = lineIndex;
    while (rowEnd + 1 < record.lines.length && !record.lines[rowEnd]!.rowEnd)
      rowEnd++;
    const row = record.lines.slice(lineIndex, rowEnd + 1);
    const unit = [...headers, ...row].map((line) => ({ line, spacing: 0 }));
    if (!simulateUnit(unit, page, footnotes)) {
      warnTableHeaderRelaxed(record);
      return;
    }
    for (const header of headers) {
      const result = attemptBodyLine(header, 0)!;
      page = result.page;
      footnotes = result.footnotes;
      warnings.push(...result.diagnostics);
    }
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
      target.visual += line.visualCount;
      target.counted += line.countedCount;
      target.bodyLines.push(line);
      enqueue(
        state,
        line.footnoteRefs,
        sectionIndex?.deepestOwnerByBlock.get(line.block) ?? 0,
      );
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
      if (sectionIndex && page.sectionTouches) {
        const owner = sectionIndex.deepestOwnerByBlock.get(record.block) ?? 0;
        for (const index of sectionIndex.sections[owner]!.ancestors) {
          page.sectionTouches.add(index);
        }
      }
      if (record.block.sectionBreak?.kind === "continuous") {
        priorAfter = 0;
        continue;
      }
      commitPage();
      priorAfter = 0;
      continue;
    }

    page = placeKeepUnit({
      record,
      blockIndex,
      bodyBlocks,
      page,
      footnotes,
      priorAfter,
      trackSections: sectionIndex !== undefined,
      simulateUnit,
      unitFromBlocks,
      emptyPage,
      commitPage: commitTargetPage,
      hasContent,
    });

    page = applyWidowOrphan({
      phase: "before",
      record,
      blockIndex,
      page,
      footnotes,
      priorAfter,
      profile,
      simulateUnit,
      unitFromBlocks,
      commitPage: commitTargetPage,
      hasContent,
    });

    const beforeLine: Snapshot[] = [];
    let lineIndex = 0;
    while (lineIndex < record.lines.length) {
      const currentLine = record.lines[lineIndex]!;
      if (record.block.kind === "table" && currentLine.rowStart) {
        let rowEnd = lineIndex;
        while (
          rowEnd + 1 < record.lines.length &&
          !record.lines[rowEnd]!.rowEnd
        )
          rowEnd++;
        const currentRow = record.lines
          .slice(lineIndex, rowEnd + 1)
          .map((line) => ({ line, spacing: 0 }));
        if (
          simulateUnit(
            currentRow,
            emptyPage(sectionIndex !== undefined),
            footnotes,
          ) &&
          !simulateUnit(currentRow, page, footnotes) &&
          hasContent(page)
        ) {
          commitPage();
          repeatTableHeader(record, lineIndex);
        }
      }
      beforeLine[lineIndex] = snapshot();
      const spacing =
        lineIndex === 0 && page.bodyLines.length > 0
          ? Math.max(priorAfter, record.style.beforeTwips)
          : 0;
      const result = attemptBodyLine(record.lines[lineIndex]!, spacing);
      if (!result) {
        const widowOrphanResult = applyWidowOrphan({
          phase: "overflow",
          record,
          page,
          lineIndex,
          beforeLine,
          profile,
          restore,
          commitRestored,
        });
        if (widowOrphanResult) {
          pages = widowOrphanResult.pages;
          page = widowOrphanResult.page;
          footnotes = widowOrphanResult.footnotes;
          lineIndex = widowOrphanResult.lineIndex;
          continue;
        }
        commitPage();
        repeatTableHeader(record, lineIndex);
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
  const diagnostics = buildDiagnostics(
    pages,
    document,
    profile,
    fonts,
    sectionIndex,
    warnings,
    styleFor,
    candidateWidth,
    sourceRangesFor,
  );

  const last = pages.at(-1)!;
  return {
    pageCount: pages.length,
    equivalentPages: pages.length - 1 + occupied(last) / usableHeight,
    totalVisualLines: pages.reduce(
      (total, placedPage) => total + placedPage.visual,
      0,
    ),
    visualLinesByPage: pages.map((placedPage) => placedPage.visual),
    countedLinesByPage: pages.map((placedPage) => placedPage.counted),
    lastPage: {
      visualLines: last.visual,
      usedTwips: occupied(last),
      usableTwips: usableHeight,
      bodyLineEquivalentsUsed: occupied(last) / bodyPitch,
      bodyLineCapacity: Math.floor(usableHeight / bodyPitch),
    },
    paragraphs: diagnostics.paragraphResults,
    warnings: diagnostics.warnings,
    ...(diagnostics.sectionResults
      ? { sections: diagnostics.sectionResults }
      : {}),
  };
}
