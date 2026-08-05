import { AgentDocxError } from "../types.js";
import {
  blockBookmark,
  isBlockId,
  type BlockId,
} from "../legal/model.js";
import type { AttachmentManifest } from "./contracts.js";
import {
  decodeDocxXml,
  parseDocxXml,
  sha256Hex,
} from "./package.js";
import { validAttachmentManifest } from "./attachments.js";
import { emittedBookmarkName, fromBookmarkName } from "./tracked.js";

export type SemanticRevisionMapEntry = {
  changeId: string;
  attribution: {
    author: { name: string; email?: string } | null;
    createdAt: string | null;
    sourceRevisionId?: string;
  };
  blockId?: BlockId | null;
  baseText?: string;
  headText?: string;
};

export type SemanticManifest = {
  schemaVersion: 1;
  generator: "agent-docx";
  mode: "clean" | "redline";
  projectId: string;
  documentId: string;
  source: string;
  sourceSha256: `sha256:${string}`;
  document: Record<string, unknown>;
  blocks: readonly {
    id: BlockId;
    bookmark: string;
    parentId: BlockId | null;
    depth: number;
    order: number;
    kind: string;
    authorities?: readonly {
      run: number;
      id: string;
      category: "cases" | "statutes" | "rules" | "other";
      short: string;
    }[];
  }[];
  emittedBlocks: readonly { bookmark: string; index: number }[];
  attachments: AttachmentManifest | null;
  revision: `sha256:${string}` | null;
  baseRevision: `sha256:${string}` | null;
  validation: Record<string, unknown>;
  dependencies: readonly {
    key: string;
    sha256: `sha256:${string}`;
    mediaType: string;
    byteLength: number;
  }[];
  revisionMap: readonly SemanticRevisionMapEntry[];
  commentMap: readonly {
    annotationId: string;
    blockWide: boolean;
    authorEmail: string | null;
  }[];
};

const unsupported = (message: string): never => {
  throw new AgentDocxError("DOCX_IMPORT_UNSUPPORTED", message);
};

const asObject = (value: unknown, label: string): Record<string, unknown> => {
  if (value === null || typeof value !== "object" || Array.isArray(value))
    unsupported(`${label} must be an object`);
  return value as Record<string, unknown>;
};

const exactKeys = (
  value: Record<string, unknown>,
  keys: readonly string[],
  label: string,
): void => {
  if (
    Object.keys(value).length !== keys.length ||
    Object.keys(value).some((key) => !keys.includes(key))
  )
    unsupported(`${label} has an unsupported property`);
};
const readManifestPayload = (bytes: Uint8Array): Record<string, unknown> => {
  let root = false;
  let payload = false;
  let payloadText = "";
  parseDocxXml(
    decodeDocxXml(bytes),
    (tag) => {
      if (tag.local === "agent-docx") root = true;
      if (root && tag.local === "payload") payload = true;
    },
    (tag) => {
      if (tag.local === "payload") payload = false;
      if (tag.local === "agent-docx") root = false;
    },
    (text) => {
      if (payload) payloadText += text;
    },
  );
  if (payloadText.length === 0)
    unsupported("Semantic manifest payload is missing");
  let parsed: unknown;
  try {
    parsed = JSON.parse(payloadText);
  } catch {
    unsupported("Semantic manifest payload is not JSON");
  }
  const manifest = asObject(parsed, "Semantic manifest");
  exactKeys(
    manifest,
    [
      "schemaVersion",
      "generator",
      "mode",
      "projectId",
      "documentId",
      "source",
      "sourceSha256",
      "document",
      "blocks",
      "attachments",
      "emittedBlocks",
      "revision",
      "baseRevision",
      "validation",
      "dependencies",
      "revisionMap",
      "commentMap",
    ],
    "Semantic manifest",
  );
  if (
    manifest.schemaVersion !== 1 ||
    manifest.generator !== "agent-docx" ||
    (manifest.mode !== "clean" && manifest.mode !== "redline") ||
    typeof manifest.projectId !== "string" ||
    typeof manifest.documentId !== "string" ||
    typeof manifest.source !== "string" ||
    typeof manifest.sourceSha256 !== "string" ||
    !/^sha256:[a-f0-9]{64}$/.test(manifest.sourceSha256) ||
    (manifest.revision !== null &&
      (typeof manifest.revision !== "string" ||
        !/^sha256:[a-f0-9]{64}$/.test(manifest.revision))) ||
    (manifest.baseRevision !== null &&
      (typeof manifest.baseRevision !== "string" ||
        !/^sha256:[a-f0-9]{64}$/.test(manifest.baseRevision))) ||
    manifest.validation === null ||
    typeof manifest.validation !== "object" ||
    Array.isArray(manifest.validation) ||
    !Array.isArray(manifest.dependencies) ||
    !Array.isArray(manifest.revisionMap) ||
    !Array.isArray(manifest.commentMap) ||
    sha256Hex(manifest.source) !== manifest.sourceSha256 ||
    !Array.isArray(manifest.blocks) ||
    !Array.isArray(manifest.emittedBlocks) ||
    (manifest.attachments !== null &&
      !validAttachmentManifest(manifest.attachments))
  )
    unsupported("Semantic manifest has an invalid version-1 shape");
  return manifest;
};

export const parseSemanticManifest = (bytes: Uint8Array): SemanticManifest => {
  const manifest = readManifestPayload(bytes);
  const rawBlocks = manifest.blocks;
  const rawEmittedBlocks = manifest.emittedBlocks;
  if (!Array.isArray(rawBlocks) || !Array.isArray(rawEmittedBlocks))
    unsupported("Semantic manifest has invalid block lists");
  const blocks: SemanticManifest["blocks"][number][] = (
    rawBlocks as unknown[]
  ).map(parseBlockRecord);
  validateBlockGraph(blocks);
  const emittedBlocks: SemanticManifest["emittedBlocks"][number][] = (
    rawEmittedBlocks as unknown[]
  ).map(parseEmittedBlockRecord);
  validateEmittedBlockGraph(emittedBlocks, blocks);
  const dependencies = parseDependencyRecords(manifest.dependencies);
  const revisionMap = parseRevisionRecords(manifest.revisionMap);
  const commentMap = parseCommentRecords(manifest.commentMap);
  const mode = manifest.mode as "clean" | "redline";
  const revision = manifest.revision as `sha256:${string}` | null;
  const baseRevision = manifest.baseRevision as `sha256:${string}` | null;
  if (
    (mode === "clean" &&
      (baseRevision !== null ||
        revisionMap.length !== 0 ||
        commentMap.length !== 0)) ||
    (mode === "redline" &&
      (revision === null || baseRevision === null || revision === baseRevision))
  )
    unsupported("Semantic manifest revision mode is inconsistent");
  return {
    schemaVersion: 1,
    generator: "agent-docx",
    mode: manifest.mode as "clean" | "redline",
    projectId: manifest.projectId as string,
    documentId: manifest.documentId as string,
    sourceSha256: manifest.sourceSha256 as `sha256:${string}`,
    source: manifest.source as string,
    document: asObject(manifest.document, "Semantic manifest document"),
    blocks,
    emittedBlocks,
    attachments: manifest.attachments as AttachmentManifest | null,
    revision: manifest.revision as `sha256:${string}` | null,
    baseRevision: manifest.baseRevision as `sha256:${string}` | null,
    validation: asObject(manifest.validation, "Semantic manifest validation"),
    dependencies,
    revisionMap,
    commentMap,
  };
};
 
const validateBlockGraph = (
  blocks: readonly SemanticManifest["blocks"][number][],
): void => {
  const blockIds = new Set<string>();
  const orders = new Set<number>();
  const validKinds = new Set([
    "paragraph",
    "heading",
    "blockquote",
    "numbered-paragraph",
    "list",
    "table",
    "caption",
    "toc",
    "toa",
    "signature",
    "certificate",
    "exhibit",
    "length-exclusion",
    "image",
    "pagebreak",
    "thematic-break",
    "sectionbreak",
  ]);
  for (const block of blocks) {
    if (
      blockIds.has(block.id) ||
      orders.has(block.order) ||
      !validKinds.has(block.kind) ||
      (block.parentId === null && block.depth !== 0)
    )
      unsupported("Semantic manifest has an invalid block graph");
    blockIds.add(block.id);
    orders.add(block.order);
  }
  if (
    [...orders]
      .sort((left, right) => left - right)
      .some((order, expected) => order !== expected)
  )
    unsupported("Semantic manifest block orders are not contiguous");
  const blocksById = new Map(blocks.map((block) => [block.id, block]));
  for (const block of blocks) {
    if (block.parentId === null) continue;
    const parent = blocksById.get(block.parentId);
    if (
      !parent ||
      parent.order >= block.order ||
      block.depth !== parent.depth + 1
    )
      unsupported("Semantic manifest block parent graph is invalid");
  }
  const expectedBookmarks = new Set(blocks.map((block) => block.bookmark));
  if (expectedBookmarks.size !== blocks.length)
    unsupported("Semantic manifest block bookmarks are not unique");
};
 
const validateEmittedBlockGraph = (
  emittedBlocks: readonly SemanticManifest["emittedBlocks"][number][],
  blocks: readonly SemanticManifest["blocks"][number][],
): void => {
  const blockIds = new Set(blocks.map((block) => block.id));
  const bookmarks = new Set<string>();
  const indexes = new Set<number>();
  for (const emitted of emittedBlocks) {
    const blockId = fromBookmarkName(emitted.bookmark);
    if (
      bookmarks.has(emitted.bookmark) ||
      indexes.has(emitted.index) ||
      (blockId !== null && !blockIds.has(blockId))
    )
      unsupported("Semantic manifest has invalid emitted block identities");
    bookmarks.add(emitted.bookmark);
    indexes.add(emitted.index);
  }
  if (
    [...indexes]
      .sort((left, right) => left - right)
      .some((index, expected) => index !== expected)
  )
    unsupported("Semantic manifest emitted block indexes are not contiguous");
};
 
const parseDependencyRecords = (
  raw: unknown,
): SemanticManifest["dependencies"] => {
  const dependencyKeys = new Set<string>();
  return (raw as unknown[]).map((entry) => {
    const dependency = asObject(entry, "Semantic manifest dependency");
    exactKeys(
      dependency,
      ["key", "sha256", "mediaType", "byteLength"],
      "Semantic manifest dependency",
    );
    const key = dependency.key;
    const digest = dependency.sha256;
    const mediaType = dependency.mediaType;
    const byteLength = dependency.byteLength;
    if (
      typeof key !== "string" ||
      key.length === 0 ||
      dependencyKeys.has(key) ||
      typeof digest !== "string" ||
      !/^sha256:[a-f0-9]{64}$/.test(digest) ||
      typeof mediaType !== "string" ||
      mediaType.length === 0 ||
      typeof byteLength !== "number" ||
      !Number.isSafeInteger(byteLength) ||
      byteLength < 0
    )
      unsupported("Semantic manifest dependency is invalid");
    const resolvedKey = key as string;
    const resolvedMediaType = mediaType as string;
    const resolvedByteLength = byteLength as number;
    dependencyKeys.add(resolvedKey);
    return {
      key: resolvedKey,
      sha256: digest as `sha256:${string}`,
      mediaType: resolvedMediaType,
      byteLength: resolvedByteLength,
    };
  });
};
 
const parseRevisionRecords = (
  raw: unknown,
): SemanticManifest["revisionMap"] => {
  const changeIds = new Set<string>();
  return (raw as unknown[]).map((entry) => {
    const revision = asObject(entry, "Semantic manifest revision map entry");
    const hasBlockId = "blockId" in revision;
    const hasBaseText = "baseText" in revision;
    const hasHeadText = "headText" in revision;
    const hasExtendedText = hasBlockId && hasBaseText && hasHeadText;
    if ((hasBlockId || hasBaseText || hasHeadText) !== hasExtendedText)
      unsupported(
        "Semantic manifest revision map entry must include blockId, baseText, and headText together",
      );
    exactKeys(
      revision,
      [
        "changeId",
        "attribution",
        ...(hasExtendedText ? ["blockId", "baseText", "headText"] : []),
      ],
      "Semantic manifest revision map entry",
    );
    const changeId = revision.changeId;
    const attribution = asObject(
      revision.attribution,
      "Semantic manifest revision attribution",
    );
    exactKeys(
      attribution,
      [
        "author",
        "createdAt",
        ...(attribution.sourceRevisionId === undefined
          ? []
          : ["sourceRevisionId"]),
      ],
      "Semantic manifest revision attribution",
    );
    const author =
      attribution.author === null
        ? null
        : asObject(attribution.author, "Semantic manifest revision author");
    const blockId = revision.blockId;
    const baseText = revision.baseText;
    const headText = revision.headText;
    if (
      (hasExtendedText &&
        ((blockId !== null &&
          (typeof blockId !== "string" || !isBlockId(blockId))) ||
          typeof baseText !== "string" ||
          typeof headText !== "string")) ||
      (!hasExtendedText &&
        (blockId !== undefined ||
          baseText !== undefined ||
          headText !== undefined))
    )
      unsupported("Semantic manifest revision map text fields are invalid");
    if (
      typeof changeId !== "string" ||
      !/^c_[a-zA-Z0-9_-]+$/.test(changeId) ||
      changeIds.has(changeId) ||
      (author !== null &&
        (typeof author.name !== "string" ||
          author.name.length === 0 ||
          (author.email !== undefined && typeof author.email !== "string"))) ||
      (attribution.createdAt !== null &&
        typeof attribution.createdAt !== "string") ||
      (attribution.sourceRevisionId !== undefined &&
        typeof attribution.sourceRevisionId !== "string")
    )
      unsupported("Semantic manifest revision map entry is invalid");
    if (author !== null)
      exactKeys(
        author,
        ["name", ...(author.email === undefined ? [] : ["email"])],
        "Semantic manifest revision author",
      );
    const safeChangeId = changeId as string;
    changeIds.add(safeChangeId);
    return {
      changeId: safeChangeId,
      attribution: {
        author:
          author === null
            ? null
            : {
                name: author.name as string,
                ...(author.email === undefined
                  ? {}
                  : { email: author.email as string }),
              },
        createdAt: attribution.createdAt as string | null,
        ...(attribution.sourceRevisionId === undefined
          ? {}
          : { sourceRevisionId: attribution.sourceRevisionId as string }),
      },
      ...(hasExtendedText
        ? {
            blockId: blockId as BlockId | null,
            baseText: baseText as string,
            headText: headText as string,
          }
        : {}),
    };
  });
};
 
const parseCommentRecords = (
  raw: unknown,
): SemanticManifest["commentMap"] => {
  const annotationIds = new Set<string>();
  return (raw as unknown[]).map((entry) => {
    const comment = asObject(entry, "Semantic manifest comment map entry");
    exactKeys(
      comment,
      ["annotationId", "blockWide", "authorEmail"],
      "Semantic manifest comment map entry",
    );
    if (
      typeof comment.annotationId !== "string" ||
      !/^a_[a-zA-Z0-9_-]+$/.test(comment.annotationId) ||
      annotationIds.has(comment.annotationId) ||
      typeof comment.blockWide !== "boolean" ||
      (comment.authorEmail !== null && typeof comment.authorEmail !== "string")
    )
      unsupported("Semantic manifest comment map entry is invalid");
    const safeAnnotationId = comment.annotationId as string;
    const safeBlockWide = comment.blockWide as boolean;
    annotationIds.add(safeAnnotationId);
    return {
      annotationId: safeAnnotationId,
      blockWide: safeBlockWide,
      authorEmail: comment.authorEmail as string | null,
    };
  });
};
 
const parseBlockRecord = (
  entry: unknown,
): SemanticManifest["blocks"][number] => {
  const block = asObject(entry, "Semantic manifest block");
  if (
    Object.keys(block).some(
      (key) =>
        ![
          "id",
          "bookmark",
          "parentId",
          "depth",
          "order",
          "kind",
          "authorities",
        ].includes(key),
    )
  )
    unsupported("Semantic manifest block has an unsupported property");
  const id = block.id;
  const bookmark = block.bookmark;
  const parentId = block.parentId;
  const depth = block.depth;
  const order = block.order;
  const kind = block.kind;
  if (
    typeof id !== "string" ||
    !isBlockId(id) ||
    typeof bookmark !== "string" ||
    bookmark !== blockBookmark(id) ||
    (parentId !== null &&
      (typeof parentId !== "string" || !isBlockId(parentId))) ||
    !Number.isSafeInteger(depth) ||
    (depth as number) < 0 ||
    !Number.isSafeInteger(order) ||
    (order as number) < 0 ||
    typeof kind !== "string"
  )
    unsupported("Semantic manifest block is invalid");
  let authorities:
    | SemanticManifest["blocks"][number]["authorities"]
    | undefined;
  if (block.authorities !== undefined) {
    if (!Array.isArray(block.authorities))
      unsupported("Semantic manifest block authorities is invalid");
    const entries = (block.authorities as unknown[]).map(
      (authority, authorityIndex) => {
        const record = asObject(
          authority,
          `Semantic manifest block authorities[${authorityIndex}]`,
        );
        exactKeys(
          record,
          ["run", "id", "category", "short"],
          "Semantic manifest block authority",
        );
        const run = record.run;
        const authorityId = record.id;
        const category = record.category;
        const short = record.short;
        if (
          !Number.isSafeInteger(run) ||
          (run as number) < 0 ||
          typeof authorityId !== "string" ||
          authorityId.length === 0 ||
          (category !== "cases" &&
            category !== "statutes" &&
            category !== "rules" &&
            category !== "other") ||
          typeof short !== "string" ||
          short.length === 0
        )
          unsupported("Semantic manifest block authority is invalid");
        return {
          run: run as number,
          id: authorityId as string,
          category: category as "cases" | "statutes" | "rules" | "other",
          short: short as string,
        };
      },
    );
    authorities = entries;
  }
  return {
    id: id as BlockId,
    bookmark: bookmark as string,
    parentId: parentId as BlockId | null,
    depth: depth as number,
    order: order as number,
    kind: kind as string,
    ...(authorities ? { authorities } : {}),
  };
};
 
const parseEmittedBlockRecord = (
  entry: unknown,
): SemanticManifest["emittedBlocks"][number] => {
  const emitted = asObject(entry, "Semantic manifest emitted block");
  exactKeys(
    emitted,
    ["bookmark", "index"],
    "Semantic manifest emitted block",
  );
  const bookmark = emitted.bookmark;
  const index = emitted.index;
  if (
    typeof bookmark !== "string" ||
    emittedBookmarkName(bookmark) === null ||
    !Number.isSafeInteger(index) ||
    (index as number) < 0
  )
    unsupported("Semantic manifest emitted block is invalid");
  return { bookmark: bookmark as string, index: index as number };
};
