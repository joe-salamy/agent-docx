import { constants as fsConstants } from "node:fs";
import { lstat, open, type FileHandle } from "node:fs/promises";
import { AgentDocxError } from "../types.js";
import {
  blockBookmark,
  emptyLitigationMetadata,
  type BlockId,
  type DocumentChrome,
  type LegalDocument,
  type LitigationMetadata,
  type ReviewAnnotation,
} from "../legal/model.js";
import { parseLegalMarkdown } from "../legal/parse.js";
import { visibleTextForBlock } from "../legal/visible-text.js";
import type {
  DocxFidelityItem,
  DocxImportResult,
  ImportAttachmentBundle,
} from "./contracts.js";
import {
  decodeDocxXml,
  readDocxParts,
  resolveOpcTarget,
  sha256Hex,
} from "./package.js";
import {
  parseRelationships,
  relationshipPartFor,
  sourcePartForRelationshipPart,
} from "./opc.js";
import { asObject, codePointCompare, unsupported } from "./helpers.js";
import {
  parseSemanticManifest,
  type SemanticManifest,
  type SemanticRevisionMapEntry,
} from "./manifest.js";
import type { ImportedAsset, AttachmentResolution } from "./attachments.js";
import {
  attachmentInventory,
  resolveAttachmentBundle,
  sourceAssetsForSemanticDocument,
} from "./attachments.js";
import {
  escapedMarkdown,
  fromBookmarkName,
  parseParagraphs,
  parseTrackedParagraphs,
  type Paragraph,
  type TrackedMaterial,
  type TrackedParagraph,
} from "./tracked.js";
import {
  cleanAnnotations,
  extendedRevisionEntry,
  parseNativeComments,
  reconstructTrackedMaterial,
  redlineAnnotations,
  resolveRevisionGroups,
  semanticBlocks,
  semanticDeletedSourceRanges,
  sourceWithVisibleBlocks,
  trackedAnnotations,
  type RedlineDecisionMap,
} from "./annotations.js";

export { parseSemanticManifest };
export type { SemanticManifest, SemanticRevisionMapEntry };
export type { TrackedParagraph, TrackedMaterial };
export type { ImportedAsset, AttachmentResolution };

const validatePackageRelationships = (
  parts: ReadonlyMap<string, Uint8Array>,
): void => {
  for (const [part, bytes] of [...parts.entries()].sort(([left], [right]) =>
    codePointCompare(left, right),
  )) {
    if (!part.endsWith(".rels")) continue;
    const sourcePart = sourcePartForRelationshipPart(part);
    if (sourcePart !== "" && !parts.has(sourcePart))
      unsupported(`Relationship part has no source part: ${part}`);
    for (const relationship of parseRelationships(
      decodeDocxXml(bytes),
      sourcePart,
    )) {
      if (relationship.external) continue;
      const target = resolveOpcTarget(sourcePart, relationship.target);
      if (!parts.has(target))
        unsupported(`Relationship target is dangling: ${part} -> ${target}`);
    }
  }
};

const sourceTextByBlockId = (
  document: LegalDocument,
): ReadonlyMap<BlockId, string> => {
  const text = new Map<BlockId, string>();
  const visit = (blocks: readonly LegalDocument["blocks"][number][]): void => {
    for (const block of blocks) {
      if ("runs" in block) text.set(block.id, visibleTextForBlock(block));
      if (block.kind === "exhibit" || block.kind === "length-exclusion")
        visit(block.blocks);
    }
  };
  visit(document.blocks);
  return text;
};

const semanticDocumentFidelity = (
  semantic: SemanticManifest | null,
  paragraphs: readonly Paragraph[],
  document: LegalDocument,
  mainPart: string,
): readonly DocxFidelityItem<"unsupported">[] => {
  if (semantic === null) return [];
  const expected = [...semantic.emittedBlocks].sort(
    (left, right) => left.index - right.index,
  );
  const actual = paragraphs.flatMap((paragraph) =>
    paragraph.bookmarkName === null ? [] : [paragraph],
  );
  const expectedNames = expected.map((entry) => entry.bookmark);
  const actualNames = actual.map((paragraph) => paragraph.bookmarkName!);
  const fidelity: DocxFidelityItem<"unsupported">[] = [];
  if (
    expectedNames.length !== actualNames.length ||
    expectedNames.some((name, index) => name !== actualNames[index])
  )
    fidelity.push({
      status: "unsupported",
      partPath: mainPart,
      relationshipId: null,
      ooxmlKind: "w:bookmarkStart",
      count: Math.max(expectedNames.length, actualNames.length),
      blockIds: expectedNames.flatMap((name) => {
        const blockId = fromBookmarkName(name);
        return blockId === null ? [] : [blockId];
      }),
      sourcePositions: [],
      explanation:
        "The emitted agent-docx bookmark sequence does not match the semantic manifest.",
    });
  const expectedBlockIds = new Set(
    expectedNames.flatMap((name) => {
      const blockId = fromBookmarkName(name);
      return blockId === null ? [] : [blockId];
    }),
  );
  const sourceText = sourceTextByBlockId(document);
  const changed = actual.flatMap((paragraph) => {
    if (
      paragraph.bookmark === null ||
      !expectedBlockIds.has(paragraph.bookmark)
    )
      return [];
    const expectedText = sourceText.get(paragraph.bookmark);
    return expectedText === undefined || expectedText === paragraph.text
      ? []
      : [paragraph.bookmark];
  });
  if (changed.length > 0)
    fidelity.push({
      status: "unsupported",
      partPath: mainPart,
      relationshipId: null,
      ooxmlKind: "w:t",
      count: changed.length,
      blockIds: changed,
      sourcePositions: [],
      explanation:
        "Visible paragraph text does not match the source preserved by the semantic manifest.",
    });
  return fidelity;
};

const requirePackage = (
  parts: ReadonlyMap<string, Uint8Array>,
): {
  mainPart: string;
  mainXml: string;
  commentsPart: string | null;
  semanticRelationshipId: string | null;
  semantic: SemanticManifest | null;
} => {
  const content = parts.get("[Content_Types].xml");
  const rootRels = parts.get("_rels/.rels");
  if (!content) unsupported("DOCX content types part is missing");
  if (!rootRels) unsupported("DOCX root relationships part is missing");
  const contentXml = decodeDocxXml(content as Uint8Array);
  if (/macroEnabled|vbaProject|encryptedPackage/i.test(contentXml))
    unsupported("Macros or encrypted packages are unsupported");
  for (const name of parts.keys())
    if (/\/(?:embeddings|activeX)\/|vbaProject|oleObject|altChunk/i.test(name))
      unsupported("Embedded executable or alternate content is unsupported");
  validatePackageRelationships(parts);
  const root = parseRelationships(decodeDocxXml(rootRels as Uint8Array), "");
  const office = root.filter((entry) => /\/officeDocument$/.test(entry.type));
  if (office.length !== 1 || office[0]!.external)
    unsupported("DOCX has no unique internal main document");
  const mainPart = resolveOpcTarget("", office[0]!.target);
  const main = parts.get(mainPart);
  if (!main) unsupported("Main document relationship is dangling");
  const rels = parts.get(relationshipPartFor(mainPart));
  const mainRelationships = rels
    ? parseRelationships(decodeDocxXml(rels), mainPart)
    : [];
  const semanticRelationships = mainRelationships.filter(
    (entry) =>
      entry.type === "https://agent-docx.dev/relationships/semantic-manifest",
  );
  if (semanticRelationships.length > 1)
    unsupported("DOCX contains more than one semantic-manifest relationship");
  const commentRelationships = mainRelationships.filter((entry) =>
    /\/comments$/.test(entry.type),
  );
  if (commentRelationships.length > 1)
    unsupported("DOCX contains more than one comments relationship");
  const commentsPart =
    commentRelationships.length === 0
      ? null
      : (() => {
          const relation = commentRelationships[0]!;
          if (relation.external)
            unsupported("DOCX comments relationship must be internal");
          const target = resolveOpcTarget(mainPart, relation.target);
          if (!parts.has(target))
            unsupported("DOCX comments relationship is dangling");
          return target;
        })();
  let semantic: SemanticManifest | null = null;
  if (semanticRelationships.length === 1) {
    const relation = semanticRelationships[0]!;
    if (relation.external)
      unsupported("DOCX semantic manifest must use an internal relationship");
    const target = resolveOpcTarget(mainPart, relation.target);
    const manifest = parts.get(target);
    if (!manifest)
      unsupported("DOCX semantic manifest relationship is dangling");
    semantic = parseSemanticManifest(manifest as Uint8Array);
  }
  return {
    mainPart,
    mainXml: decodeDocxXml(main as Uint8Array),
    commentsPart,
    semanticRelationshipId:
      semanticRelationships.length === 1 ? semanticRelationships[0]!.id : null,
    semantic,
  };
};

const loadInput = async (input: string | Uint8Array): Promise<Uint8Array> => {
  if (typeof input !== "string") return input;
  const entry = await lstat(input).catch(() => null);
  if (!entry || !entry.isFile() || entry.isSymbolicLink())
    throw new AgentDocxError(
      "INPUT_NOT_FOUND",
      `DOCX input is not a regular file: ${input}`,
    );
  const flags = fsConstants.O_RDONLY | (fsConstants.O_NOFOLLOW ?? 0);
  let handle: FileHandle | undefined;
  try {
    handle = await open(input, flags);
    const stats = await handle.stat();
    if (!stats.isFile() || stats.dev !== entry.dev || stats.ino !== entry.ino)
      throw new AgentDocxError(
        "INPUT_NOT_FOUND",
        `DOCX input changed while opening: ${input}`,
      );
    return await handle.readFile();
  } catch (error) {
    if (error instanceof AgentDocxError) throw error;
    const cause = error as NodeJS.ErrnoException;
    if (
      cause.code === "ENOENT" ||
      cause.code === "ENOTDIR" ||
      cause.code === "ELOOP"
    )
      throw new AgentDocxError(
        "INPUT_NOT_FOUND",
        `DOCX input not found: ${input}`,
      );
    throw new AgentDocxError(
      "INTERNAL_ERROR",
      error instanceof Error ? error.message : String(error),
    );
  } finally {
    await handle?.close().catch(() => undefined);
  }
};
const loadSemanticPackage = async (
  input: string | Uint8Array,
  attachments: ImportAttachmentBundle | undefined,
  requiredMode?: "redline",
) => {
  const bytes = await loadInput(input);
  const parts = await readDocxParts(bytes);
  const packageInfo = requirePackage(parts);
  if (requiredMode !== undefined) {
    const semantic = packageInfo.semantic;
    if (!semantic)
      unsupported(
        "Redline resolution requires an agent-docx semantic manifest",
      );
    if ((semantic as SemanticManifest).mode !== requiredMode)
      unsupported("Redline resolution requires a redline semantic manifest");
  }
  const resolvedAttachments = await resolveAttachmentBundle(
    packageInfo.semantic?.attachments ?? null,
    attachments,
  );
  const sourceAssets = sourceAssetsForSemanticDocument(
    packageInfo.semantic,
    resolvedAttachments,
    parts,
  );
  return {
    parts,
    packageInfo,
    attachments: resolvedAttachments,
    sourceAssets,
  };
};

export type InspectedDocxMaterial = {
  source: string;
  semantic: SemanticManifest | null;
  assets: Readonly<Record<string, ImportedAsset>>;
  tracked: TrackedMaterial | null;
  result: Extract<DocxImportResult, { inspectOnly: true }>;
};

/** Reads a DOCX into the closed, fidelity-reported inspect-only contract. */
export const inspectDocxMaterial = async (
  input: string | Uint8Array,
  _options: { attachments?: ImportAttachmentBundle } = {},
): Promise<InspectedDocxMaterial> => {
  const { parts, packageInfo, attachments, sourceAssets } =
    await loadSemanticPackage(input, _options.attachments);
  const { mainPart, mainXml, commentsPart, semanticRelationshipId, semantic } =
    packageInfo;
  const trackedParsed =
    semantic?.mode === "redline"
      ? parseTrackedParagraphs(mainXml, mainPart)
      : null;
  const tracked = trackedParsed
    ? reconstructTrackedMaterial(
        semantic!,
        trackedParsed.paragraphs,
        sourceAssets.assets,
      )
    : null;
  const parsed =
    tracked === null
      ? parseParagraphs(mainXml, mainPart)
      : {
          paragraphs: tracked.paragraphs.map((paragraph) => ({
            bookmark: paragraph.bookmark,
            bookmarkName:
              paragraph.bookmark === null
                ? null
                : blockBookmark(paragraph.bookmark),
            text: paragraph.headText,
            heading: paragraph.heading,
            sourcePart: paragraph.sourcePart,
            comments: paragraph.comments,
          })),
          unsupported: trackedParsed!.unsupported,
        };
  const markdown = parsed.paragraphs
    .map((paragraph) => {
      const marker = paragraph.bookmark
        ? `<!-- agent-docx:block id="${paragraph.bookmark}" -->\n`
        : "";
      const prefix = paragraph.heading
        ? `${"#".repeat(paragraph.heading)} `
        : "";
      return `${marker}${prefix}${escapedMarkdown(paragraph.text)}`;
    })
    .join("\n\n")
    .concat(parsed.paragraphs.length === 0 ? "" : "\n");
  const source = tracked?.headSource ?? semantic?.source ?? markdown;
  const semanticDocument = semantic
    ? asObject(semantic.document, "Semantic manifest document")
    : null;
  const documentId = semantic?.documentId ?? "imported";
  const document = parseLegalMarkdown(source, {
    projectId: semantic?.projectId ?? "standalone",
    documentId,
    metadata: (semanticDocument?.metadata ??
      emptyLitigationMetadata()) as LitigationMetadata,
    chrome: (semanticDocument?.chrome ?? {}) as DocumentChrome,
    assets: sourceAssets.assets,
    requireMarkers: semantic !== null,
  }).document;
  const nativeComments = parseNativeComments(
    commentsPart === null ? undefined : decodeDocxXml(parts.get(commentsPart)!),
  );
  const cleanCommentAnnotations =
    semantic?.mode === "clean" && nativeComments.size > 0
      ? cleanAnnotations(parsed.paragraphs, document, nativeComments)
      : [];
  const semanticFidelity: DocxFidelityItem<"preserved" | "unsupported">[] =
    semantic
      ? [
          {
            status: "preserved",
            partPath: "customXml/itemAgentDocx.xml",
            relationshipId: semanticRelationshipId,
            ooxmlKind: "agent-docx:semantic-manifest",
            count: 1,
            blockIds: [],
            sourcePositions: [],
            explanation:
              "The agent-docx semantic manifest preserves the original source and block identity.",
          },
          ...semanticDocumentFidelity(
            semantic,
            parsed.paragraphs,
            document,
            mainPart,
          ),
        ]
      : [];
  const attachmentFidelity: DocxFidelityItem<"externalized" | "unsupported">[] =
    semantic?.attachments
      ? attachments.complete
        ? semantic.attachments.entries.map((entry) => ({
            status: "externalized" as const,
            partPath: "customXml/itemAgentDocx.xml",
            relationshipId: semanticRelationshipId,
            ooxmlKind: "agent-docx:attachment",
            count: 1,
            blockIds: [],
            sourcePositions: [],
            explanation: `External attachment ${entry.name} was hash-verified from the authorized bundle.`,
          }))
        : [
            {
              status: "unsupported" as const,
              partPath: "customXml/itemAgentDocx.xml",
              relationshipId: semanticRelationshipId,
              ooxmlKind: "agent-docx:attachment",
              count: semantic.attachments.entries.length,
              blockIds: [],
              sourcePositions: [],
              explanation:
                "DOCX declares external attachments but no authorized complete attachment bundle was supplied.",
            },
          ]
      : [];
  const missingAssetFidelity: DocxFidelityItem<"unsupported">[] =
    sourceAssets.unresolved.map((name) => ({
      status: "unsupported",
      partPath: "customXml/itemAgentDocx.xml",
      relationshipId: semanticRelationshipId,
      ooxmlKind: "agent-docx:asset",
      count: 1,
      blockIds: [],
      sourcePositions: [],
      explanation: `Semantic asset ${name} cannot be matched to an embedded or authorized external payload.`,
    }));
  const unexpectedCommentFidelity: DocxFidelityItem<"unsupported">[] =
    tracked === null &&
    cleanCommentAnnotations.length === 0 &&
    nativeComments.size > 0
      ? [
          {
            status: "unsupported",
            partPath: commentsPart ?? mainPart,
            relationshipId: null,
            ooxmlKind: "w:comment",
            count: nativeComments.size,
            blockIds: [],
            sourcePositions: [],
            explanation:
              "Comments without a tracked agent-docx semantic manifest cannot be anchored safely.",
          },
        ]
      : [];
  const fidelity: DocxFidelityItem<
    "preserved" | "normalized" | "externalized" | "unsupported"
  >[] = [
    ...parsed.unsupported,
    ...semanticFidelity,
    ...attachmentFidelity,
    ...missingAssetFidelity,
    ...unexpectedCommentFidelity,
    ...parsed.paragraphs.map((paragraph) => ({
      status: "normalized" as const,
      partPath: mainPart,
      relationshipId: null,
      ooxmlKind: "w:p",
      count: 1,
      blockIds: paragraph.bookmark ? [paragraph.bookmark] : [],
      sourcePositions: [],
      explanation: paragraph.bookmark
        ? "Bookmark identity is retained; Word paragraph formatting is normalized to Markdown semantics."
        : "Unbookmarked paragraph is represented with a newly assigned Markdown block identity.",
    })),
  ];
  const unsupported = fidelity.some((item) => item.status === "unsupported");
  const annotations: readonly ReviewAnnotation[] =
    tracked !== null
      ? trackedAnnotations(semantic!, tracked, document, nativeComments)
      : cleanCommentAnnotations;
  return {
    source,
    semantic,
    assets: sourceAssets.assets,
    tracked,
    result: {
      schemaVersion: 1,
      inspectOnly: true,
      mode: "inspect",
      output: null,
      sourceSha256: unsupported ? null : sha256Hex(source),
      baseRevision: null,
      headRevision: null,
      revisions: [],
      recognized: {
        blocks: document.blocks,
        footnotes: document.footnotes,
        annotations,
        assets: {
          ...attachmentInventory(parts),
          ...attachments.inventory,
        },
      },
      fidelity: {
        overall: unsupported ? "unsupported" : "normalized",
        items: fidelity,
      },
    },
  };
};

export const inspectDocx = async (
  input: string | Uint8Array,
  options: { attachments?: ImportAttachmentBundle } = {},
): Promise<Extract<DocxImportResult, { inspectOnly: true }>> =>
  (await inspectDocxMaterial(input, options)).result;
export type RedlineInspection = {
  semantic: SemanticManifest;
  tracked: TrackedMaterial;
  resolution: "none" | "complete";
  rejectedSource: string | null;
  annotations: readonly ReviewAnnotation[];
  readonly decisions: RedlineDecisionMap;
};

export const inspectRedlineResolution = async (
  input: string | Uint8Array,
  options: { attachments?: ImportAttachmentBundle } = {},
): Promise<RedlineInspection> => {
  const { parts, packageInfo, sourceAssets } = await loadSemanticPackage(
    input,
    options.attachments,
    "redline",
  );
  const semantic = packageInfo.semantic as SemanticManifest;
  if (sourceAssets.unresolved.length > 0)
    unsupported(
      `Semantic asset cannot be resolved: ${sourceAssets.unresolved.join(", ")}`,
    );
  const { paragraphs } = parseTrackedParagraphs(
    packageInfo.mainXml,
    packageInfo.mainPart,
  );
  const trackedComments = parseNativeComments(
    packageInfo.commentsPart === null
      ? undefined
      : decodeDocxXml(parts.get(packageInfo.commentsPart)!),
  );
  const projectedDocument = asObject(
    semantic.document,
    "Semantic manifest document",
  );
  const document = parseLegalMarkdown(semantic.source, {
    projectId: semantic.projectId,
    documentId: semantic.documentId,
    metadata: projectedDocument.metadata as LitigationMetadata,
    chrome: (projectedDocument.chrome ?? {}) as DocumentChrome,
    assets: sourceAssets.assets,
    requireMarkers: true,
  }).document;
  const tracked: TrackedMaterial = {
    baseSource: semantic.source,
    headSource: semantic.source,
    paragraphs,
  };
  const actualRevisionCount = paragraphs.reduce(
    (total, paragraph) => total + paragraph.revisions.length,
    0,
  );
  if (semantic.revisionMap.every(extendedRevisionEntry)) {
    const expectedRevisionCount = semantic.revisionMap.reduce(
      (total, entry) =>
        total + (entry.baseText === "" || entry.headText === "" ? 1 : 2),
      0,
    );
    if (
      actualRevisionCount > 0 &&
      actualRevisionCount !== expectedRevisionCount
    )
      unsupported(
        "Redline contains a partially resolved set of tracked changes",
      );
  }
  if (actualRevisionCount > 0) {
    const reconstructed = reconstructTrackedMaterial(
      semantic,
      paragraphs,
      sourceAssets.assets,
    );
    const visibleByBlock = new Map<BlockId, string>();
    for (const paragraph of paragraphs) {
      if (!paragraph.bookmark) continue;
      if (visibleByBlock.has(paragraph.bookmark))
        unsupported("Redline has duplicate agent-docx block bookmarks");
      visibleByBlock.set(paragraph.bookmark, paragraph.visibleText);
    }
    const annotations = redlineAnnotations(
      semantic,
      reconstructed,
      document,
      trackedComments,
      visibleByBlock,
    );
    return {
      semantic,
      tracked: reconstructed,
      resolution: "none",
      rejectedSource: null,
      annotations,
      decisions: {},
    };
  }
  const resolved = resolveRevisionGroups(semantic, document, tracked);
  const blocks = semanticBlocks(document);
  const annotations = redlineAnnotations(
    semantic,
    tracked,
    document,
    trackedComments,
    resolved.visibleByBlock,
  );
  if (semantic.revisionMap.length === 0)
    return {
      semantic,
      tracked,
      resolution: "none",
      rejectedSource: null,
      annotations,
      decisions: {},
    };
  const rejectedSource = sourceWithVisibleBlocks(
    semantic.source,
    blocks,
    resolved.visibleByBlock,
    semanticDeletedSourceRanges(semantic),
  );
  const baseSource =
    resolved.baseByBlock.size === 0
      ? semantic.source
      : sourceWithVisibleBlocks(
          semantic.source,
          blocks,
          resolved.baseByBlock,
          semanticDeletedSourceRanges(semantic),
        );
  return {
    semantic,
    tracked: {
      baseSource,
      headSource: semantic.source,
      paragraphs,
    },
    resolution: "complete",
    rejectedSource,
    annotations,
    decisions: resolved.decisions,
  };
};
