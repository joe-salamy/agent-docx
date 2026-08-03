import type {
  FlowBlock,
  FootnoteDefinition as FlowFootnoteDefinition,
  NormalizedDocument,
  TableCell as FlowTableCell,
  TextFlowBlock,
} from "../markdown.js";
import type { SourcePosition } from "../types.js";
import type {
  FootnoteDefinition,
  InlineParagraph,
  InlineRun,
  LegalBlock,
  LegalDocument,
  LegalListBlock,
} from "./model.js";

type SemanticFlowHints = {
  legalKind?: LegalBlock["kind"] | "footnote";
  listOrdered?: boolean;
  listLevel?: number;
  numberedLevel?: number;
  sequence?: string;
};

type IdentifiedTextFlowBlock = TextFlowBlock &
  SemanticFlowHints & {
    legalBlockId?: string;
    image?: {
      source: string;
      alt: string;
      widthTwips: number;
      heightTwips: number;
    };
  };

type SectionBreakFlowBlock = FlowBlock & {
  sectionBreak?: {
    kind: "next-page" | "continuous";
    pageNumber?: { format: "decimal" | "lower-roman" | "upper-roman"; start: number };
    legalBlockId: string;
  };
};

const asTextFlowBlock = (
  kind: TextFlowBlock["kind"],
  paragraph:
    | InlineParagraph
    | (Pick<LegalBlock, "position" | "segments"> & {
        runs: readonly InlineRun[];
      }),
  level?: number,
  legalBlockId?: string,
  semantic: SemanticFlowHints = {},
): IdentifiedTextFlowBlock => {
  const runs = paragraph.runs.map((run) => ({
    text: `${run.text}${run.hardBreakAfter ? "\n" : ""}`,
    bold: run.bold,
    italic: run.italic,
    ...(run.literal ? { literal: true } : {}),
    ...(run.footnoteId ? { footnoteId: run.footnoteId } : {}),
    ...(run.link ? { link: run.link } : {}),
    ...(run.referenceTarget ? { referenceTarget: run.referenceTarget } : {}),
    ...(run.strikethrough ? { strikethrough: true } : {}),
  }));
  return {
    kind,
    runs,
    normalizedText: paragraph.runs
      .map((run) => `${run.text}${run.hardBreakAfter ? "\n" : ""}`)
      .join(""),
    sourceSegments: [...paragraph.segments],
    position: paragraph.position,
    ...(level === undefined ? {} : { level }),
    ...(legalBlockId === undefined ? {} : { legalBlockId }),
    footnoteRefs: paragraph.runs.flatMap((run) =>
      run.footnoteId ? [run.footnoteId] : [],
    ),
    ...semantic,
  };
};

const mergedCell = (paragraphs: readonly InlineParagraph[]): FlowTableCell => {
  const position: SourcePosition = paragraphs[0]?.position ?? {
    start: { line: 1, column: 1, offset: 0 },
    end: { line: 1, column: 1, offset: 0 },
  };
  const mergedRuns = paragraphs.flatMap((paragraph, index) => [
    ...paragraph.runs,
    ...(index + 1 < paragraphs.length
      ? [
          {
            text: "\n",
            bold: false,
            italic: false,
            strikethrough: false,
            literal: false,
            hardBreakAfter: false,
          } satisfies InlineRun,
        ]
      : []),
  ]);
  const segments = paragraphs.flatMap((paragraph) => paragraph.segments);
  return {
    runs: mergedRuns.map((run) => ({
      text: `${run.text}${run.hardBreakAfter ? "\n" : ""}`,
      bold: run.bold,
      italic: run.italic,
      ...(run.literal ? { literal: true } : {}),
      ...(run.footnoteId ? { footnoteId: run.footnoteId } : {}),
      ...(run.link ? { link: run.link } : {}),
      ...(run.referenceTarget ? { referenceTarget: run.referenceTarget } : {}),
      ...(run.strikethrough ? { strikethrough: true } : {}),
    })),
    normalizedText: mergedRuns
      .map((run) => `${run.text}${run.hardBreakAfter ? "\n" : ""}`)
      .join(""),
    sourceSegments: segments,
    position,
    alignment: null,
  };
};

const appendList = (
  list: LegalListBlock,
  blocks: FlowBlock[],
  paragraphs: TextFlowBlock[],
): void => {
  for (const item of list.items) {
    for (const paragraph of item.paragraphs) {
      const flow = asTextFlowBlock("list", paragraph, undefined, undefined, {
        legalKind: "list",
        listOrdered: list.ordered,
        listLevel: Math.max(0, list.depth - 1),
      });
      blocks.push(flow);
      paragraphs.push(flow);
    }
    for (const child of item.children) appendList(child, blocks, paragraphs);
  }
};

const appendBlock = (
  block: LegalBlock,
  document: LegalDocument,
  blocks: FlowBlock[],
  paragraphs: TextFlowBlock[],
): void => {
  if (block.kind === "paragraph") {
    const flow = asTextFlowBlock("paragraph", block, undefined, block.id, {
      legalKind: "paragraph",
    });
    blocks.push(flow);
    paragraphs.push(flow);
    return;
  }
  if (block.kind === "heading") {
    const flow = asTextFlowBlock("heading", block, block.level, block.id, {
      legalKind: "heading",
    });
    blocks.push(flow);
    paragraphs.push(flow);
    return;
  }
  if (block.kind === "blockquote") {
    const flow = asTextFlowBlock("blockquote", block, undefined, block.id, {
      legalKind: "blockquote",
    });
    blocks.push(flow);
    paragraphs.push(flow);
    return;
  }
  if (block.kind === "numbered-paragraph") {
    const flow = asTextFlowBlock("list", block, undefined, block.id, {
      legalKind: "numbered-paragraph",
      numberedLevel: block.level - 1,
      sequence: block.sequence,
    });
    blocks.push(flow);
    paragraphs.push(flow);
    return;
  }
  if (block.kind === "list") {
    appendList(block, blocks, paragraphs);
    return;
  }
  if (block.kind === "table") {
    blocks.push({
      kind: "table",
      position: block.position,
      alignments: block.align,
      rows: block.rows.map((row) => row.map((cell) => mergedCell(cell.paragraphs))),
      legalBlockId: block.id,
      legalKind: "table",
    } as FlowBlock);
    return;
  }
  if (block.kind === "thematic-break") {
    blocks.push({ kind: "thematic-break", position: block.position });
    return;
  }
  if (block.kind === "pagebreak" || block.kind === "sectionbreak") {
    blocks.push({
      kind: "pagebreak",
      position: block.position,
      ...(block.kind === "sectionbreak"
        ? {
            sectionBreak: {
              kind: block.breakKind,
              ...(block.pageNumber ? { pageNumber: block.pageNumber } : {}),
              legalBlockId: block.id,
            },
          }
        : {}),
    } as SectionBreakFlowBlock);
    return;
  }
  if (block.kind === "exhibit") {
    if (blocks.length > 0)
      blocks.push({ kind: "pagebreak", position: block.position });
    const cover = asTextFlowBlock(
      "paragraph",
      {
        position: block.position,
        segments: block.segments,
        runs: [
          {
            text: `${block.label}\nExternal attachment: ${block.source}`,
            bold: true,
            italic: false,
            strikethrough: false,
            literal: false,
            hardBreakAfter: false,
          },
        ],
      },
      undefined,
      block.id,
      { legalKind: "exhibit" },
    );
    blocks.push(cover);
    paragraphs.push(cover);
    for (const child of block.blocks) appendBlock(child, document, blocks, paragraphs);
    return;
  }
  if (block.kind === "length-exclusion") {
    for (const child of block.blocks) appendBlock(child, document, blocks, paragraphs);
    return;
  }
  if (block.kind === "caption") {
    const flow = asTextFlowBlock(
      "paragraph",
      {
        position: block.position,
        segments: block.segments,
        runs: [
          {
            text: `${document.metadata.court}\n${document.metadata.caseName}\n${document.metadata.docketNumber}\n${document.metadata.documentTitle}`,
            bold: true,
            italic: false,
            strikethrough: false,
            literal: false,
            hardBreakAfter: false,
          },
        ],
      },
      undefined,
      block.id,
      { legalKind: "caption" },
    );
    blocks.push(flow);
    paragraphs.push(flow);
    return;
  }
  if (block.kind === "signature") {
    const counsel = document.metadata.counsel.find((entry) => entry.id === block.counselId);
    const flow = asTextFlowBlock(
      "paragraph",
      {
        position: block.position,
        segments: block.segments,
        runs: [
          {
            text: counsel?.name ?? block.counselId,
            bold: false,
            italic: false,
            strikethrough: false,
            literal: false,
            hardBreakAfter: false,
          },
        ],
      },
      undefined,
      block.id,
      { legalKind: "signature" },
    );
    blocks.push(flow);
    paragraphs.push(flow);
    return;
  }
  if (block.kind === "certificate") {
    const certificate = document.metadata.certificates.find(
      (entry) => entry.id === block.certificateId,
    );
    const flow = asTextFlowBlock(
      "paragraph",
      {
        position: block.position,
        segments: block.segments,
        runs: [
          {
            text:
              certificate?.kind === "service"
                ? certificate.statement
                : certificate?.kind === "compliance"
                  ? `Certificate of compliance (${certificate.basis})`
                  : block.certificateId,
            bold: false,
            italic: false,
            strikethrough: false,
            literal: false,
            hardBreakAfter: false,
          },
        ],
      },
      undefined,
      block.id,
      { legalKind: "certificate" },
    );
    blocks.push(flow);
    paragraphs.push(flow);
    return;
  }
  if (block.kind === "image") {
    const flow = asTextFlowBlock(
      "paragraph",
      {
        position: block.position,
        segments: block.segments,
        runs: [
          {
            text: block.alt,
            bold: false,
            italic: false,
            strikethrough: false,
            literal: false,
            hardBreakAfter: false,
          },
        ],
      },
      undefined,
      block.id,
      { legalKind: "image" },
    );
    flow.image = {
      source: block.source,
      alt: block.alt,
      widthTwips: block.widthTwips,
      heightTwips: block.heightTwips,
    };
    blocks.push(flow);
    paragraphs.push(flow);
    return;
  }
  if (block.kind === "toc" || block.kind === "toa") {
    const flow = asTextFlowBlock(
      "heading",
      {
        position: block.position,
        segments: block.segments,
        runs: [
          {
            text: block.kind === "toc" ? "Table of Contents" : "Table of Authorities",
            bold: true,
            italic: false,
            strikethrough: false,
            literal: false,
            hardBreakAfter: false,
          },
        ],
      },
      1,
      block.id,
      { legalKind: block.kind },
    );
    blocks.push(flow);
    paragraphs.push(flow);
  }
};

const lowerFootnote = (footnote: FootnoteDefinition): FlowFootnoteDefinition => {
  const lowered = footnote.paragraphs.map((paragraph) =>
    asTextFlowBlock("footnote", paragraph, undefined, footnote.id, {
      legalKind: "footnote",
    }),
  );
  return {
    id: footnote.label,
    position: footnote.position,
    blocks: lowered,
    footnoteRefs: lowered.flatMap((block) => block.footnoteRefs),
  };
};

export const lowerLegalDocument = (document: LegalDocument): NormalizedDocument => {
  const blocks: FlowBlock[] = [];
  const paragraphs: TextFlowBlock[] = [];
  for (const block of document.blocks) appendBlock(block, document, blocks, paragraphs);
  return {
    blocks,
    footnotes: new Map(
      document.footnotes.map((footnote) => [footnote.label, lowerFootnote(footnote)]),
    ),
    paragraphs,
  };
};
