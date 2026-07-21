import {
  Document,
  FootnoteReferenceRun,
  Packer,
  PageBreak,
  Paragraph,
  TextRun,
  type ISpacingProperties,
  type ParagraphChild,
} from "docx";
import type { FlowBlock, NormalizedDocument } from "../markdown.js";
import type { LayoutProfile, TextStyle } from "../types.js";

function spacing(style: TextStyle): ISpacingProperties {
  return style.lineSpacing.rule === "auto"
    ? {
        line: style.lineSpacing.numerator,
        lineRule: "auto",
        before: style.beforeTwips,
        after: style.afterTwips,
      }
    : {
        line: style.lineSpacing.twips,
        lineRule: style.lineSpacing.rule === "exact" ? "exact" : "atLeast",
        before: style.beforeTwips,
        after: style.afterTwips,
      };
}

function paragraphOptions(style: TextStyle, widowControl: boolean) {
  return {
    keepNext: style.keepWithNext,
    keepLines: style.keepLines,
    widowControl,
    spacing: spacing(style),
    indent: {
      left: style.leftIndentTwips,
      right: style.rightIndentTwips,
      firstLine: style.firstLineIndentTwips,
      hanging: style.hangingIndentTwips,
    },
  };
}

function textChildren(
  block: FlowBlock,
  style: TextStyle,
  fontFamily: string,
  footnoteIds: ReadonlyMap<string, number>,
): ParagraphChild[] {
  return block.runs.map((run) => {
    if (run.footnoteId !== undefined)
      return new FootnoteReferenceRun(footnoteIds.get(run.footnoteId)!);
    return new TextRun({
      text: run.text,
      font: fontFamily,
      size: style.fontSizeTwips / 10,
      bold: style.bold || run.bold,
      italics: style.italic || run.italic,
      break: run.text === "\n" ? 1 : undefined,
    });
  });
}

export async function generateDocx(
  flow: NormalizedDocument,
  profile: LayoutProfile,
): Promise<Uint8Array> {
  const footnoteIds = new Map<string, number>();
  let nextId = 1;
  for (const id of flow.footnotes.keys()) footnoteIds.set(id, nextId++);
  const footnotes: Record<string, { children: readonly Paragraph[] }> = {};
  for (const [id, block] of flow.footnotes) {
    const numeric = footnoteIds.get(id)!;
    footnotes[String(numeric)] = {
      children: [
        new Paragraph({
          ...paragraphOptions(
            profile.footnote,
            profile.pagination.widowOrphanControl,
          ),
          children: textChildren(
            block,
            profile.footnote,
            profile.requestedFontFamily,
            footnoteIds,
          ),
        }),
      ],
    };
  }
  const children: Paragraph[] = [];
  for (const block of flow.blocks) {
    if (block.kind === "pagebreak") {
      children.push(new Paragraph({ children: [new PageBreak()] }));
      continue;
    }
    const style =
      block.kind === "heading"
        ? profile.headings[String(block.level ?? 1) as "1"]
        : block.kind === "blockquote"
          ? profile.blockquote
          : block.kind === "list"
            ? profile.list
            : profile.body;
    children.push(
      new Paragraph({
        ...paragraphOptions(style, profile.pagination.widowOrphanControl),
        children: textChildren(
          block,
          style,
          profile.requestedFontFamily,
          footnoteIds,
        ),
      }),
    );
  }
  if (children.length === 0) children.push(new Paragraph({ children: [] }));
  const document = new Document({
    footnotes,
    sections: [
      {
        properties: {
          page: {
            size: {
              width: profile.page.widthTwips,
              height: profile.page.heightTwips,
            },
            margin: {
              top: profile.page.marginsTwips.top,
              right: profile.page.marginsTwips.right,
              bottom: profile.page.marginsTwips.bottom,
              left: profile.page.marginsTwips.left,
              header: profile.page.headerTwips,
              footer: profile.page.footerTwips,
              gutter: profile.page.gutterTwips,
            },
          },
        },
        children,
      },
    ],
  });
  return Packer.toBuffer(document);
}
