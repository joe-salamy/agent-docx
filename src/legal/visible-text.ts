import type { AddressableBlock, InlineRun } from "./model.js";

export const visibleRuns = (runs: readonly InlineRun[]): string =>
  runs.map((run) => `${run.text}${run.hardBreakAfter ? "\n" : ""}`).join("");

export const visibleBlock = (block: AddressableBlock): string => {
  if (block.kind === "footnote")
    return block.paragraphs
      .map((paragraph) => visibleRuns(paragraph.runs))
      .join("\n");
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
