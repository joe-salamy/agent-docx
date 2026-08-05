import type { TextFlowBlock } from "../markdown.js";
import type { LayoutProfile, TextStyle } from "./profile.js";

/**
 * One style policy for flow blocks, shared by the deterministic paginator,
 * the DOCX renderer, and redline export so every path applies the same
 * paragraph style and style ID to the same block.
 */
export const flowStyleFor = (
  block: TextFlowBlock,
  profile: LayoutProfile,
): { style: TextStyle; styleId: string } => {
  if (block.legalKind === "caption")
    return { styleId: "AgentDocxCaption", style: { ...profile.body, bold: true } };
  if (block.legalKind === "signature")
    return { styleId: "AgentDocxSignature", style: profile.body };
  if (block.legalKind === "certificate")
    return { styleId: "AgentDocxCertificate", style: profile.body };
  if (block.legalKind === "toc")
    return { styleId: "AgentDocxTOCHeading", style: profile.headings["1"] };
  if (block.legalKind === "toa")
    return { styleId: "AgentDocxTOAHeading", style: profile.headings["1"] };
  if (block.kind === "heading") {
    const level = Math.min(6, Math.max(1, block.level ?? 1));
    return {
      styleId: `AgentDocxHeading${level}`,
      style: profile.headings[String(level) as "1"],
    };
  }
  if (block.kind === "blockquote")
    return { styleId: "AgentDocxBlockQuote", style: profile.blockquote };
  if (block.kind === "list" && block.legalKind === "numbered-paragraph") {
    const level = Math.min(3, Math.max(0, block.numberedLevel ?? 0));
    return { styleId: `AgentDocxNumbered${level + 1}`, style: profile.list };
  }
  if (block.kind === "list")
    return { styleId: "AgentDocxList", style: profile.list };
  if (block.kind === "footnote")
    return { styleId: "AgentDocxFootnote", style: profile.footnote };
  return { styleId: "AgentDocxBody", style: profile.body };
};
