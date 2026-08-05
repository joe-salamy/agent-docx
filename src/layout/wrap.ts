import type {
  FlowBlock,
  NormalizedDocument,
  SectionIndex,
  TableFlowBlock,
  TextFlowBlock,
} from "../markdown.js";
import type { LoadedFonts } from "../resolve.js";
import type { Diagnostic, SourcePosition } from "../types.js";
import type {
  LayoutProfile,
  TextStyle,
} from "./profile.js";
import type {
  ParagraphDiagnostic,
  SectionDiagnostic,
} from "../measurement.js";

export type WrappedLine = {
  used: number;
  available: number;
  height: number;
  text: string;
  block: FlowBlock;
  counted: boolean;
  visualCount: number;
  countedCount: number;
  start: number;
  end: number;
  contentEnd: number;
  startCause: "soft" | "hard" | "start";
  overflowed: boolean;
  footnoteRefs: string[];
  rowStart?: boolean;
  rowEnd?: boolean;
};

export type PlacedFootnoteLine = {
  footnoteId: string;
  line: WrappedLine;
  ownerIndex: number;
};

export type Page = {
  bodyUsed: number;
  footnoteUsed: number;
  visual: number;
  counted: number;
  bodyLines: WrappedLine[];
  footnoteLines: PlacedFootnoteLine[];
  sectionTouches?: Set<number>;
};

export type PendingFootnote = {
  footnoteId: string;
  blocks: readonly (readonly WrappedLine[])[];
  blockIndex: number;
  nextLine: number;
  warningEmitted: boolean;
  ownerIndex: number;
};

export type FootnoteState = {
  placed: Set<string>;
  pending: PendingFootnote[];
  relaxed: Set<string>;
};

export type WrappedBlock = {
  block: FlowBlock;
  style: TextStyle;
  lines: WrappedLine[];
  headerLineCount?: number;
};

export type WrappedTable = {
  lines: WrappedLine[];
  headerLineCount: number;
  gridWidths: number[];
};

export type Snapshot = {
  pages: Page[];
  page: Page;
  footnotes: FootnoteState;
};

export type PaginationState = {
  pages: Page[];
  page: Page;
  footnotes: FootnoteState;
};

export type WrapText = (
  block: TextFlowBlock,
  style: TextStyle,
  available: number,
  fonts: LoadedFonts,
  warnings: Diagnostic[],
  lineCapExclusions: LayoutProfile["pagination"]["lineCapExclusions"],
) => WrappedLine[];

export type WrapTable = (
  block: TableFlowBlock,
  profile: LayoutProfile,
  fonts: LoadedFonts,
  usableWidth: number,
  usableHeight: number,
  warnings: Diagnostic[],
) => WrappedTable;

export type StyleFor = (
  profile: LayoutProfile,
  block: TextFlowBlock,
) => TextStyle;

export type NaturalHeight = (
  fonts: LoadedFonts,
  block: TextFlowBlock,
  style: TextStyle,
  start: number,
  end: number,
) => number;

export type LinePitch = (natural: number, style: TextStyle) => number;

export function wrapBodyBlocks(
  document: NormalizedDocument,
  profile: LayoutProfile,
  fonts: LoadedFonts,
  usableWidth: number,
  usableHeight: number,
  warnings: Diagnostic[],
  wrap: WrapText,
  wrapTable: WrapTable,
  styleFor: StyleFor,
  naturalHeight: NaturalHeight,
  linePitch: LinePitch,
): { bodyBlocks: WrappedBlock[]; bodyPitch: number } {
  const metricProbe: TextFlowBlock = {
    kind: "paragraph",
    runs: [{ text: "Ag", bold: false, italic: false }],
    normalizedText: "Ag",
    sourceSegments: [],
    position: document.blocks[0]!.position,
    footnoteRefs: [],
  };
  const bodyPitch = linePitch(
    naturalHeight(fonts, metricProbe, profile.body, 0, 2),
    profile.body,
  );
  const bodyBlocks: WrappedBlock[] = document.blocks.map((block) => {
    if (block.kind === "pagebreak")
      return { block, style: profile.body, lines: [] };
    if (block.kind === "thematic-break") {
      const style: TextStyle = {
        ...profile.body,
        beforeTwips: profile.thematicBreak.beforeTwips,
        afterTwips: profile.thematicBreak.afterTwips,
        keepWithNext: profile.thematicBreak.keepWithNext,
        keepLines: true,
        lineSpacing: {
          rule: "exact",
          twips: profile.thematicBreak.thicknessTwips,
        },
      };
      return {
        block,
        style,
        lines: [
          {
            used: usableWidth,
            available: usableWidth,
            height: profile.thematicBreak.thicknessTwips,
            text: "",
            block,
            counted: false,
            visualCount: 0,
            countedCount: 0,
            start: 0,
            end: 0,
            contentEnd: 0,
            startCause: "start",
            overflowed: false,
            footnoteRefs: [],
          },
        ],
      };
    }
    if (block.kind === "table") {
      const wrapped = wrapTable(
        block,
        profile,
        fonts,
        usableWidth,
        usableHeight,
        warnings,
      );
      return {
        block,
        style: {
          ...profile.table.body,
          beforeTwips: 0,
          afterTwips: 0,
          keepWithNext: false,
          keepLines: false,
        },
        lines: wrapped.lines,
        headerLineCount: wrapped.headerLineCount,
      };
    }
    if (block.image) {
      if (block.image.widthTwips > usableWidth)
        warnings.push({
          code: "PROFILE_CONSTRAINT_VIOLATION",
          severity: "warning",
          message: `Image ${block.image.source} exceeds body width`,
          position: block.position,
        });
      return {
        block,
        style: profile.body,
        lines: [
          {
            used: block.image.widthTwips,
            available: usableWidth,
            height: block.image.heightTwips,
            text: "",
            block,
            counted: false,
            visualCount: Math.ceil(block.image.heightTwips / bodyPitch),
            countedCount: 0,
            start: 0,
            end: 0,
            contentEnd: 0,
            startCause: "start",
            overflowed: false,
            footnoteRefs: [],
          },
        ],
      };
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
  return { bodyBlocks, bodyPitch };
}

export function wrapFootnotes(
  document: NormalizedDocument,
  profile: LayoutProfile,
  usableWidth: number,
  fonts: LoadedFonts,
  warnings: Diagnostic[],
  wrap: WrapText,
): (id: string) => readonly (readonly WrappedLine[])[] {
  const footnoteCache = new Map<string, readonly (readonly WrappedLine[])[]>();
  const wrappedFootnote = (id: string) => {
    let blocks = footnoteCache.get(id);
    if (blocks) return blocks;
    const definition = document.footnotes.get(id)!;
    const style = profile.footnote;
    blocks = definition.blocks.map((block) =>
      wrap(
        block,
        style,
        usableWidth - style.leftIndentTwips - style.rightIndentTwips,
        fonts,
        warnings,
        profile.pagination.lineCapExclusions,
      ),
    );
    footnoteCache.set(id, blocks);
    return blocks;
  };
  return wrappedFootnote;
}

export type UnitLine = { line: WrappedLine; spacing: number };
export type UnitFromBlocks = (
  indexes: readonly number[],
  terminatingLines: number | null,
  firstPriorAfter: number,
  startHasBody: boolean,
) => UnitLine[];
export type SimulateUnit = (
  unit: readonly UnitLine[],
  startPage: Page,
  startFootnotes: FootnoteState,
) => boolean;
export type CommitPage = (target: Page) => Page;
export type EmptyPage = (trackSections?: boolean) => Page;
export type HasContent = (page: Page) => boolean;

export type PlaceKeepUnitArgs = {
  record: WrappedBlock;
  blockIndex: number;
  bodyBlocks: readonly WrappedBlock[];
  page: Page;
  footnotes: FootnoteState;
  priorAfter: number;
  trackSections: boolean;
  simulateUnit: SimulateUnit;
  unitFromBlocks: UnitFromBlocks;
  emptyPage: EmptyPage;
  commitPage: CommitPage;
  hasContent: HasContent;
};

export function placeKeepUnit(args: PlaceKeepUnitArgs): Page {
  const {
    record,
    blockIndex,
    bodyBlocks,
    page,
    footnotes,
    priorAfter,
    trackSections,
    simulateUnit,
    unitFromBlocks,
    emptyPage,
    commitPage,
    hasContent,
  } = args;
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
        simulateUnit(emptyUnit, emptyPage(trackSections), footnotes) &&
        !simulateUnit(currentUnit, page, footnotes) &&
        hasContent(page)
      ) {
        return commitPage(page);
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
      simulateUnit(emptyUnit, emptyPage(trackSections), footnotes) &&
      !simulateUnit(currentUnit, page, footnotes) &&
      hasContent(page)
    ) {
      return commitPage(page);
    }
  }
  return page;
}

type WidowOrphanBeforeArgs = {
  phase: "before";
  record: WrappedBlock;
  blockIndex: number;
  page: Page;
  footnotes: FootnoteState;
  priorAfter: number;
  profile: LayoutProfile;
  simulateUnit: SimulateUnit;
  unitFromBlocks: UnitFromBlocks;
  commitPage: CommitPage;
  hasContent: HasContent;
};

type WidowOrphanOverflowArgs = {
  phase: "overflow";
  record: WrappedBlock;
  page: Page;
  lineIndex: number;
  beforeLine: readonly Snapshot[];
  profile: LayoutProfile;
  restore: (saved: Snapshot) => PaginationState;
  commitRestored: (saved: PaginationState) => PaginationState;
};

type WidowOrphanArgs = WidowOrphanBeforeArgs | WidowOrphanOverflowArgs;
export type WidowOrphanOverflowResult = PaginationState & {
  lineIndex: number;
};

export function applyWidowOrphan(args: WidowOrphanBeforeArgs): Page;
export function applyWidowOrphan(
  args: WidowOrphanOverflowArgs,
): WidowOrphanOverflowResult | null;
export function applyWidowOrphan(
  args: WidowOrphanArgs,
): Page | WidowOrphanOverflowResult | null {
  if (args.phase === "before") {
    const {
      record,
      blockIndex,
      page,
      footnotes,
      priorAfter,
      profile,
      simulateUnit,
      unitFromBlocks,
      commitPage,
      hasContent,
    } = args;
    if (
      record.block.kind !== "table" &&
      profile.pagination.widowOrphanControl &&
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
        return commitPage(page);
      }
    }
    return page;
  }

  const { record, page, lineIndex, beforeLine, profile, restore, commitRestored } =
    args;
  const remaining = record.lines.length - lineIndex;
  const placedOnPage = page.bodyLines.filter(
    (line) => line.block === record.block,
  ).length;
  if (
    record.block.kind !== "table" &&
    profile.pagination.widowOrphanControl &&
    remaining < profile.pagination.widowLines &&
    placedOnPage >= profile.pagination.orphanLines
  ) {
    const move = Math.min(profile.pagination.orphanLines, placedOnPage);
    const targetIndex = lineIndex - move;
    const restored = restore(beforeLine[targetIndex]!);
    const committed = commitRestored(restored);
    return { ...committed, lineIndex: targetIndex };
  }
  return null;
}

export type SourceRange = {
  position: SourcePosition;
  precision: "exact" | "node";
};
export type CandidateWidth = (
  block: TextFlowBlock,
  start: number,
  end: number,
  fonts: LoadedFonts,
  style: TextStyle,
) => number;
export type SourceRangesFor = (
  block: TextFlowBlock,
  start: number,
  end: number,
) => readonly SourceRange[];

export type BuildDiagnosticsResult = {
  paragraphResults: ParagraphDiagnostic[];
  sectionResults: SectionDiagnostic[] | undefined;
  warnings: Diagnostic[];
};

export function buildDiagnostics(
  pages: readonly Page[],
  document: NormalizedDocument,
  profile: LayoutProfile,
  fonts: LoadedFonts,
  sectionIndex: SectionIndex | undefined,
  warnings: Diagnostic[],
  styleFor: StyleFor,
  candidateWidth: CandidateWidth,
  sourceRangesFor: SourceRangesFor,
): BuildDiagnosticsResult {
  const paragraphResults: ParagraphDiagnostic[] = [];
  let paragraphIndex = 0;
  for (const block of document.blocks) {
    if (
      block.kind === "pagebreak" ||
      block.kind === "table" ||
      block.kind === "thematic-break"
    )
      continue;
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
      const penultimate = occurrences.at(-2)?.line ?? null;
      const style = styleFor(profile, block);
      const oneLineReduction =
        penultimate && last.line.startCause === "soft"
          ? {
              estimatedRemovalTwips: Math.max(
                0,
                candidateWidth(
                  block,
                  penultimate.start,
                  last.line.contentEnd,
                  fonts,
                  style,
                ) - penultimate.available,
              ),
              basis: "deterministic-tail-width-deficit" as const,
              confidence: "heuristic" as const,
            }
          : null;
      paragraphResults.push({
        source: "deterministic",
        index: paragraphIndex++,
        position: block.position,
        startPage: occurrences[0]!.page,
        endPage: last.page,
        visualLines: occurrences.length,
        lastLineText: last.line.text,
        lastLineTextRange: {
          start: last.line.start,
          end: last.line.contentEnd,
        },
        lastLineSourceRanges: sourceRangesFor(
          block,
          last.line.start,
          last.line.contentEnd,
        ),
        lastLineUsedTwips: last.line.used,
        lastLineAvailableTwips: last.line.available,
        lastLineUnusedTwips: last.line.available - last.line.used,
        lastLineRatio:
          last.line.available === 0 ? 0 : last.line.used / last.line.available,
        lastLineOverflow: last.line.overflowed,
        penultimateLineText: penultimate?.text ?? null,
        penultimateLineUnusedTwips: penultimate
          ? penultimate.available - penultimate.used
          : null,
        oneLineReduction,
        preview: block.normalizedText.replace(/\s+/g, " ").trim().slice(0, 80),
      });
    }
  }

  let sectionResults: SectionDiagnostic[] | undefined;
  if (sectionIndex) {
    type MutableSectionPage = {
      page: number;
      bodyVisualLines: number;
      footnoteVisualLines: number;
      visualLines: number;
      countedLines: number;
    };
    const sectionPages = sectionIndex.sections.map(
      () => new Map<number, MutableSectionPage>(),
    );
    pages.forEach((placedPage, pageIndex) => {
      const pageNumber = pageIndex + 1;
      for (const section of placedPage.sectionTouches ?? []) {
        sectionPages[section]!.set(pageNumber, {
          page: pageNumber,
          bodyVisualLines: 0,
          footnoteVisualLines: 0,
          visualLines: 0,
          countedLines: 0,
        });
      }
      for (const line of placedPage.bodyLines) {
        const owner = sectionIndex.deepestOwnerByBlock.get(line.block) ?? 0;
        for (const section of sectionIndex.sections[owner]!.ancestors) {
          let entry = sectionPages[section]!.get(pageNumber);
          if (!entry) {
            entry = {
              page: pageNumber,
              bodyVisualLines: 0,
              footnoteVisualLines: 0,
              visualLines: 0,
              countedLines: 0,
            };
            sectionPages[section]!.set(pageNumber, entry);
          }
          entry.bodyVisualLines += line.visualCount;
          entry.visualLines += line.visualCount;
          entry.countedLines += line.countedCount;
        }
      }
      for (const placed of placedPage.footnoteLines) {
        for (const section of sectionIndex.sections[placed.ownerIndex]!
          .ancestors) {
          let entry = sectionPages[section]!.get(pageNumber);
          if (!entry) {
            entry = {
              page: pageNumber,
              bodyVisualLines: 0,
              footnoteVisualLines: 0,
              visualLines: 0,
              countedLines: 0,
            };
            sectionPages[section]!.set(pageNumber, entry);
          }
          entry.footnoteVisualLines += placed.line.visualCount;
          entry.visualLines += placed.line.visualCount;
          entry.countedLines += placed.line.countedCount;
        }
      }
    });
    sectionResults = sectionIndex.sections.map((section) => {
      const diagnosticPages = [...sectionPages[section.index]!.values()].sort(
        (a, b) => a.page - b.page,
      );
      return {
        source: "deterministic",
        index: section.index,
        parentIndex: section.parentIndex,
        heading: section.heading,
        position: section.position,
        empty: section.empty,
        startPage: diagnosticPages[0]?.page ?? null,
        endPage: diagnosticPages.at(-1)?.page ?? null,
        pageCount: diagnosticPages.length,
        bodyVisualLines: diagnosticPages.reduce(
          (total, page) => total + page.bodyVisualLines,
          0,
        ),
        footnoteVisualLines: diagnosticPages.reduce(
          (total, page) => total + page.footnoteVisualLines,
          0,
        ),
        visualLines: diagnosticPages.reduce(
          (total, page) => total + page.visualLines,
          0,
        ),
        countedLines: diagnosticPages.reduce(
          (total, page) => total + page.countedLines,
          0,
        ),
        pages: diagnosticPages,
      };
    });
  }

  return { paragraphResults, sectionResults, warnings };
}
