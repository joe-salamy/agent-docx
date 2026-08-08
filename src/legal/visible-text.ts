import { AgentDocxError } from "../types.js";
import type {
  AddressableBlock,
  InlineRun,
  LitigationMetadata,
} from "./model.js";
export const visibleRuns = (runs: readonly InlineRun[]): string =>
  runs.map((run) => `${run.text}${run.hardBreakAfter ? "\n" : ""}`).join("");

export const visibleBlock = (
  block: AddressableBlock,
  metadata?: LitigationMetadata,
  depth = 0,
): string => {
  if (depth > 100)
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      "Legal block nesting exceeds 100 levels",
    );
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
        ...item.children.map((child) =>
          visibleBlock(child, metadata, depth + 1),
        ),
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
    return block.blocks
      .map((child) => visibleBlock(child, metadata, depth + 1))
      .join("\n");
  if (block.kind === "image") return block.alt;
  if (block.kind === "signature")
    return (
      metadata?.counsel.find((entry) => entry.id === block.counselId)?.name ??
      block.counselId
    );
  if (block.kind === "certificate") {
    const certificate = metadata?.certificates.find(
      (entry) => entry.id === block.certificateId,
    );
    return certificate?.kind === "service"
      ? certificate.statement
      : certificate?.kind === "compliance"
        ? `Certificate of compliance (${certificate.basis})`
        : block.certificateId;
  }
  return "";
};

export const visibleTextForBlock = (
  block: AddressableBlock,
  metadata?: LitigationMetadata,
): string => visibleBlock(block, metadata);
