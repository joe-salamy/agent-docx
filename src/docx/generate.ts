import {
  AlignmentType,
  Bookmark,
  BorderStyle,
  Document,
  FootnoteReferenceRun,
  Packer,
  PageBreak,
  Paragraph,
  Table,
  TableCell,
  TableLayoutType,
  TableRow,
  TextRun,
  WidthType,
  type ISpacingProperties,
  type ParagraphChild,
} from "docx";
import type {
  NormalizedDocument,
  TableFlowBlock,
  TextFlowBlock,
} from "../markdown.js";
import type { LayoutProfile, SourcePosition, TextStyle } from "../types.js";

type TextContent = Pick<TextFlowBlock, "runs">;
export type BodyParagraphManifestEntry = {
  id: string;
  index: number;
  position: SourcePosition;
  preview: string;
};
export type GeneratedDocx = {
  bytes: Uint8Array;
  bodyParagraphs: readonly BodyParagraphManifestEntry[];
};

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
  block: TextContent,
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

function tableGridWidths(
  block: TableFlowBlock,
  profile: LayoutProfile,
  usableWidth: number,
): number[] {
  const columns = block.rows[0]?.length ?? 0;
  const fixed =
    profile.table.cellPaddingTwips.left +
    profile.table.cellPaddingTwips.right +
    2 * profile.table.borderTwips;
  const styleFloor = (style: TextStyle) =>
    Math.max(0, style.leftIndentTwips) +
    Math.max(0, style.rightIndentTwips) +
    Math.max(0, style.firstLineIndentTwips - style.hangingIndentTwips);
  const floor =
    fixed +
    Math.max(styleFloor(profile.table.header), styleFloor(profile.table.body)) +
    1;
  const widths = Array<number>(columns).fill(floor);
  let remaining =
    usableWidth - widths.reduce((total, value) => total + value, 0);
  for (let index = 0; remaining > 0; index = (index + 1) % columns) {
    widths[index]!++;
    remaining--;
  }
  return widths;
}

export async function generateDocx(
  flow: NormalizedDocument,
  profile: LayoutProfile,
): Promise<GeneratedDocx> {
  const footnoteIds = new Map<string, number>();
  let nextId = 1;
  for (const id of flow.footnotes.keys()) footnoteIds.set(id, nextId++);
  const footnotes: Record<string, { children: readonly Paragraph[] }> = {};
  for (const [id, definition] of flow.footnotes) {
    const numeric = footnoteIds.get(id)!;
    footnotes[String(numeric)] = {
      children: definition.blocks.map((block, blockIndex) => {
        const style: TextStyle = {
          ...profile.footnote,
          beforeTwips: blockIndex === 0 ? 0 : profile.footnote.beforeTwips,
          afterTwips:
            blockIndex === definition.blocks.length - 1
              ? 0
              : profile.footnote.afterTwips,
          keepWithNext:
            blockIndex < definition.blocks.length - 1 &&
            profile.footnote.keepWithNext,
        };
        return new Paragraph({
          ...paragraphOptions(style, profile.pagination.widowOrphanControl),
          children: textChildren(
            block,
            style,
            profile.requestedFontFamily,
            footnoteIds,
          ),
        });
      }),
    };
  }
  const usableWidth =
    profile.page.widthTwips -
    profile.page.marginsTwips.left -
    profile.page.marginsTwips.right -
    profile.page.gutterTwips;
  const children: Array<Paragraph | Table> = [];
  const borderSize =
    profile.table.borderTwips === 0
      ? 0
      : Math.max(1, Math.floor((profile.table.borderTwips * 2) / 5 + 0.5));
  const tableBorder = {
    style:
      profile.table.borderTwips === 0 ? BorderStyle.NONE : BorderStyle.SINGLE,
    size: borderSize,
    color: "000000",
  };
  const bodyParagraphs: BodyParagraphManifestEntry[] = [];
  let paragraphIndex = 0;
  for (const block of flow.blocks) {
    if (block.kind === "pagebreak") {
      children.push(new Paragraph({ children: [new PageBreak()] }));
      continue;
    }
    if (block.kind === "thematic-break") {
      children.push(
        new Paragraph({
          keepNext: profile.thematicBreak.keepWithNext,
          keepLines: true,
          widowControl: profile.pagination.widowOrphanControl,
          spacing: {
            before: profile.thematicBreak.beforeTwips,
            after: profile.thematicBreak.afterTwips,
            line: profile.thematicBreak.thicknessTwips,
            lineRule: "exact",
          },
          indent: { left: 0, right: 0, firstLine: 0, hanging: 0 },
          border: {
            bottom: {
              style: BorderStyle.SINGLE,
              color: "000000",
              size: Math.max(
                1,
                Math.floor(
                  (profile.thematicBreak.thicknessTwips * 2) / 5 + 0.5,
                ),
              ),
            },
          },
          children: [],
        }),
      );
      continue;
    }
    if (block.kind === "table") {
      const gridWidths = tableGridWidths(block, profile, usableWidth);
      children.push(
        new Table({
          width: { size: usableWidth, type: WidthType.DXA },
          columnWidths: gridWidths,
          layout: TableLayoutType.FIXED,
          borders: {
            top: tableBorder,
            bottom: tableBorder,
            left: tableBorder,
            right: tableBorder,
            insideHorizontal: tableBorder,
            insideVertical: tableBorder,
          },
          margins: {
            marginUnitType: WidthType.DXA,
            top: profile.table.cellPaddingTwips.top,
            right: profile.table.cellPaddingTwips.right,
            bottom: profile.table.cellPaddingTwips.bottom,
            left: profile.table.cellPaddingTwips.left,
          },
          rows: block.rows.map((row, rowIndex) => {
            const style =
              rowIndex === 0 ? profile.table.header : profile.table.body;
            return new TableRow({
              cantSplit: true,
              tableHeader: rowIndex === 0 && profile.table.repeatHeader,
              children: row.map(
                (cell, columnIndex) =>
                  new TableCell({
                    width: {
                      size: gridWidths[columnIndex]!,
                      type: WidthType.DXA,
                    },
                    margins: {
                      marginUnitType: WidthType.DXA,
                      top: profile.table.cellPaddingTwips.top,
                      right: profile.table.cellPaddingTwips.right,
                      bottom: profile.table.cellPaddingTwips.bottom,
                      left: profile.table.cellPaddingTwips.left,
                    },
                    borders: {
                      top: tableBorder,
                      bottom: tableBorder,
                      left: tableBorder,
                      right: tableBorder,
                    },
                    children: [
                      new Paragraph({
                        ...paragraphOptions(
                          style,
                          profile.pagination.widowOrphanControl,
                        ),
                        alignment:
                          cell.alignment === "center"
                            ? AlignmentType.CENTER
                            : cell.alignment === "right"
                              ? AlignmentType.RIGHT
                              : AlignmentType.LEFT,
                        children: textChildren(
                          cell,
                          style,
                          profile.requestedFontFamily,
                          footnoteIds,
                        ),
                      }),
                    ],
                  }),
              ),
            });
          }),
        }),
      );
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
    const id = `mpc_body_${String(paragraphIndex).padStart(6, "0")}`;
    const manifest = {
      id,
      index: paragraphIndex++,
      position: block.position,
      preview: block.normalizedText.replace(/\s+/g, " ").trim().slice(0, 80),
    };
    bodyParagraphs.push(manifest);
    const bookmark = new Bookmark({
      id,
      children: textChildren(
        block,
        style,
        profile.requestedFontFamily,
        footnoteIds,
      ),
    });
    children.push(
      new Paragraph({
        ...paragraphOptions(style, profile.pagination.widowOrphanControl),
        children: [bookmark],
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
  return {
    bytes: await Packer.toBuffer(document),
    bodyParagraphs,
  };
}
