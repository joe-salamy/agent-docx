import type { Font } from "fontkit";
import LineBreaker from "linebreak";
import type { InlineRun, TableFlowBlock, TextFlowBlock } from "../markdown.js";
import type { LoadedFonts } from "../resolve.js";
import { AgentDocxError } from "../types.js";
import type { LayoutProfile, TextStyle } from "./profile.js";

export const round = (number: number) =>
  Math.sign(number) * Math.floor(Math.abs(number) + 0.5);

export const role = (run: InlineRun, style: TextStyle) =>
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

const TAB_STOP_TWIPS = 720;

function width(font: Font, text: string, size: number, start = 0) {
  if (!text.includes("\t"))
    return round(
      (font
        .layout(text)
        .positions.reduce((sum, position) => sum + position.xAdvance, 0) *
        size) /
        font.unitsPerEm,
    );
  let total = 0;
  let segment = "";
  const appendSegment = () => {
    if (segment.length === 0) return;
    total +=
      (font
        .layout(segment)
        .positions.reduce((sum, position) => sum + position.xAdvance, 0) *
        size) /
      font.unitsPerEm;
    segment = "";
  };
  for (const character of text) {
    if (character !== "\t") {
      segment += character;
      continue;
    }
    appendSegment();
    const remainder = (start + total) % TAB_STOP_TWIPS;
    total += remainder === 0 ? TAB_STOP_TWIPS : TAB_STOP_TWIPS - remainder;
  }
  appendSegment();
  return round(total);
}

export function candidateWidth(
  block: TextFlowBlock,
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
        total,
      );
    }
    cursor = next;
  }
  return total;
}

/**
 * Allocates table column widths from content metrics. Shared by the
 * deterministic paginator and the DOCX renderer so both paths agree.
 */
export function tableColumnWidths(
  block: TableFlowBlock,
  profile: LayoutProfile,
  fonts: LoadedFonts,
  usableWidth: number,
): number[] {
  const columnCount = block.rows[0]?.length ?? 0;
  if (columnCount === 0)
    throw new AgentDocxError(
      "INVALID_LAYOUT",
      "A table must contain at least one column.",
      { position: block.position as unknown as Record<string, never> },
    );
  const styleFloor = (style: TextStyle) =>
    Math.max(0, style.leftIndentTwips) +
    Math.max(0, style.rightIndentTwips) +
    Math.max(0, style.firstLineIndentTwips - style.hangingIndentTwips);
  const fixed =
    profile.table.cellPaddingTwips.left +
    profile.table.cellPaddingTwips.right +
    2 * profile.table.borderTwips;
  const structuralFloor =
    fixed +
    Math.max(styleFloor(profile.table.header), styleFloor(profile.table.body)) +
    1;
  const widths = Array<number>(columnCount).fill(structuralFloor);
  if (widths.reduce((total, value) => total + value, 0) > usableWidth)
    throw new AgentDocxError(
      "INVALID_LAYOUT",
      "Table structural column floors exceed the usable page width.",
      { position: block.position as unknown as Record<string, never> },
    );
  const minimums = [...widths];
  const preferred = [...widths];
  for (let rowIndex = 0; rowIndex < block.rows.length; rowIndex++) {
    const style = rowIndex === 0 ? profile.table.header : profile.table.body;
    const structural = fixed + styleFloor(style);
    const row = block.rows[rowIndex]!;
    for (let columnIndex = 0; columnIndex < columnCount; columnIndex++) {
      const cell = row[columnIndex];
      if (!cell) continue;
      const cellBlock: TextFlowBlock = {
        kind: "paragraph",
        runs: cell.runs,
        normalizedText: cell.normalizedText,
        sourceSegments: cell.sourceSegments,
        position: cell.position,
        footnoteRefs: [],
      };
      let unbreakable = 0;
      let tokenStart = 0;
      const breaker = new LineBreaker(cell.normalizedText);
      let opportunity;
      while ((opportunity = breaker.nextBreak()) !== null) {
        const tokenEnd = opportunity.position;
        unbreakable = Math.max(
          unbreakable,
          candidateWidth(cellBlock, tokenStart, tokenEnd, fonts, style),
        );
        tokenStart = tokenEnd;
      }
      unbreakable = Math.max(
        unbreakable,
        candidateWidth(
          cellBlock,
          tokenStart,
          cell.normalizedText.length,
          fonts,
          style,
        ),
      );
      minimums[columnIndex] = Math.max(
        minimums[columnIndex]!,
        structural + Math.max(1, unbreakable),
      );
      preferred[columnIndex] = Math.max(
        preferred[columnIndex]!,
        structural +
          Math.max(
            1,
            candidateWidth(
              cellBlock,
              0,
              cell.normalizedText.length,
              fonts,
              style,
            ),
          ),
      );
    }
  }
  const growToward = (targets: readonly number[]) => {
    let remaining =
      usableWidth - widths.reduce((total, value) => total + value, 0);
    const deficits = targets.map((target, index) =>
      Math.max(0, target - widths[index]!),
    );
    const totalDeficit = deficits.reduce((total, value) => total + value, 0);
    if (remaining <= 0 || totalDeficit === 0) return;
    const allocation = deficits.map((deficit) =>
      Math.floor((Math.min(remaining, totalDeficit) * deficit) / totalDeficit),
    );
    const used = allocation.reduce((total, value) => total + value, 0);
    let leftover = Math.min(remaining, totalDeficit) - used;
    const order = deficits
      .map((deficit, index) => ({
        index,
        fraction:
          (Math.min(remaining, totalDeficit) * deficit) / totalDeficit -
          allocation[index]!,
      }))
      .sort((a, b) => b.fraction - a.fraction || a.index - b.index);
    for (const entry of order) {
      if (leftover-- <= 0) break;
      allocation[entry.index]!++;
    }
    for (let index = 0; index < widths.length; index++)
      widths[index]! += allocation[index]!;
  };
  growToward(minimums);
  growToward(preferred);
  let surplus = usableWidth - widths.reduce((total, value) => total + value, 0);
  for (let index = 0; surplus > 0; index = (index + 1) % widths.length) {
    widths[index]!++;
    surplus--;
  }
  return widths;
}
