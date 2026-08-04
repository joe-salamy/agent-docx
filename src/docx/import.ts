import { createHash } from "node:crypto";
import { lstat, readFile, readdir } from "node:fs/promises";
import { basename, isAbsolute, relative, resolve, sep } from "node:path";
import type { SaxesTagNS } from "saxes";
import { AgentDocxError } from "../types.js";
import {
  blockBookmark,
  emptyLitigationMetadata,
  isBlockId,
  type BlockId,
  type DocumentChrome,
  type LegalDocument,
  type LitigationMetadata,
  type ReviewAnnotation,
} from "../legal/model.js";
import { parseLegalMarkdown } from "../legal/parse.js";
import type {
  AttachmentManifest,
  DocxFidelityItem,
  DocxImportResult,
  ImportAttachmentBundle,
} from "./contracts.js";
import {
  decodeDocxXml,
  docxXmlAttribute,
  parseDocxXml,
  readDocxParts,
  resolveOpcTarget,
} from "./package.js";

const sha256 = (value: Uint8Array | string): `sha256:${string}` =>
  `sha256:${createHash("sha256").update(value).digest("hex")}`;

type Relationship = {
  id: string;
  type: string;
  target: string;
  external: boolean;
};

type Paragraph = {
  bookmark: BlockId | null;
  bookmarkName: string | null;
  text: string;
  heading: number | null;
  sourcePart: string;
  comments: readonly TrackedCommentAnchor[];
};

type NativeRevision = {
  id: string;
  kind: "ins" | "del" | "moveFrom" | "moveTo";
  author: string;
  date: string | null;
  text: string;
};

type NativeComment = {
  id: string;
  author: string;
  date: string | null;
  text: string;
};

type TrackedCommentAnchor = {
  id: string;
  start: number;
  end: number;
};

export type TrackedParagraph = {
  bookmark: BlockId | null;
  baseText: string;
  headText: string;
  visibleText: string;
  heading: number | null;
  sourcePart: string;
  revisions: readonly NativeRevision[];
  comments: readonly TrackedCommentAnchor[];
};

export type TrackedMaterial = {
  baseSource: string;
  headSource: string;
  paragraphs: readonly TrackedParagraph[];
};

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

const validAttachmentManifest = (
  value: unknown,
): value is AttachmentManifest => {
  if (value === null || typeof value !== "object" || Array.isArray(value))
    return false;
  const manifest = value as Record<string, unknown>;
  if (
    manifest.schemaVersion !== 1 ||
    !Array.isArray(manifest.entries) ||
    Object.keys(manifest).length !== 2 ||
    Object.keys(manifest).some(
      (key) => !["schemaVersion", "entries"].includes(key),
    )
  )
    return false;
  return manifest.entries.every((entry) => {
    if (entry === null || typeof entry !== "object" || Array.isArray(entry))
      return false;
    const record = entry as Record<string, unknown>;
    return (
      Object.keys(record).length === 5 &&
      ["name", "mediaType", "byteLength", "sha256", "payloadPath"].every(
        (key) => key in record,
      ) &&
      typeof record.name === "string" &&
      typeof record.mediaType === "string" &&
      Number.isSafeInteger(record.byteLength) &&
      (record.byteLength as number) >= 0 &&
      typeof record.sha256 === "string" &&
      /^sha256:[a-f0-9]{64}$/.test(record.sha256) &&
      typeof record.payloadPath === "string" &&
      record.payloadPath.startsWith("files/") &&
      !record.payloadPath.includes("..")
    );
  });
};

const parseSemanticManifest = (bytes: Uint8Array): SemanticManifest => {
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
    sha256(manifest.source) !== manifest.sourceSha256 ||
    !Array.isArray(manifest.blocks) ||
    !Array.isArray(manifest.emittedBlocks) ||
    (manifest.attachments !== null &&
      !validAttachmentManifest(manifest.attachments))
  )
    unsupported("Semantic manifest has an invalid version-1 shape");
  const rawBlocks = manifest.blocks;
  const rawEmittedBlocks = manifest.emittedBlocks;
  if (!Array.isArray(rawBlocks) || !Array.isArray(rawEmittedBlocks))
    unsupported("Semantic manifest has invalid block lists");
  const blocks: SemanticManifest["blocks"][number][] = (
    rawBlocks as unknown[]
  ).map((entry) => {
    const block = asObject(entry, "Semantic manifest block");
    exactKeys(
      block,
      ["id", "bookmark", "parentId", "depth", "order", "kind"],
      "Semantic manifest block",
    );
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
    return {
      id: id as BlockId,
      bookmark: bookmark as string,
      parentId: parentId as BlockId | null,
      depth: depth as number,
      order: order as number,
      kind: kind as string,
    };
  });
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
  const emittedBlocks: SemanticManifest["emittedBlocks"][number][] = (
    rawEmittedBlocks as unknown[]
  ).map((entry) => {
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
  });
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
  const dependencyKeys = new Set<string>();
  const dependencies: SemanticManifest["dependencies"] = (
    manifest.dependencies as unknown[]
  ).map((entry) => {
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
  const changeIds = new Set<string>();
  const revisionMap: SemanticManifest["revisionMap"] = (
    manifest.revisionMap as unknown[]
  ).map((entry) => {
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
  const annotationIds = new Set<string>();
  const commentMap: SemanticManifest["commentMap"] = (
    manifest.commentMap as unknown[]
  ).map((entry) => {
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

const relationshipPart = (part: string): string => {
  const components = part.split("/");
  const name = components.pop();
  if (!name)
    throw new AgentDocxError("DOCX_IMPORT_UNSUPPORTED", "Invalid main part");
  return [...components, "_rels", `${name}.rels`].join("/");
};

const relationships = (
  xml: string,
  sourcePart: string,
): readonly Relationship[] => {
  const found: Relationship[] = [];
  const ids = new Set<string>();
  parseDocxXml(xml, (tag) => {
    if (tag.local !== "Relationship") return;
    const id = docxXmlAttribute(tag, "Id");
    const type = docxXmlAttribute(tag, "Type");
    const target = docxXmlAttribute(tag, "Target");
    const mode = docxXmlAttribute(tag, "TargetMode");
    if (!id || !type || !target || ids.has(id))
      throw new AgentDocxError(
        "DOCX_IMPORT_UNSUPPORTED",
        "Malformed relationship",
      );
    ids.add(id);
    const external = mode === "External";
    if (external) {
      if (!/\/hyperlink$/.test(type))
        throw new AgentDocxError(
          "DOCX_IMPORT_UNSUPPORTED",
          "External non-hyperlink relationships are forbidden",
        );
      try {
        const url = new URL(target);
        if (!["http:", "https:", "mailto:"].includes(url.protocol))
          throw new Error("unsupported scheme");
      } catch {
        throw new AgentDocxError(
          "DOCX_IMPORT_UNSUPPORTED",
          "External hyperlink has an unsupported target",
        );
      }
    } else resolveOpcTarget(sourcePart, target);
    found.push({ id, type, target, external });
  });
  return found;
};

const sourcePartForRelationshipPart = (part: string): string => {
  if (part === "_rels/.rels") return "";
  const components = part.split("/");
  const relationshipName = components.pop();
  const relationshipDirectory = components.pop();
  const name =
    relationshipName?.endsWith(".rels") === true
      ? relationshipName.slice(0, -".rels".length)
      : "";
  if (relationshipDirectory !== "_rels" || name.length === 0)
    unsupported(`Malformed OPC relationship part: ${part}`);
  return [...components, name].join("/");
};

const validatePackageRelationships = (
  parts: ReadonlyMap<string, Uint8Array>,
): void => {
  for (const [part, bytes] of [...parts.entries()].sort(([left], [right]) =>
    left.localeCompare(right),
  )) {
    if (!part.endsWith(".rels")) continue;
    const sourcePart = sourcePartForRelationshipPart(part);
    if (sourcePart !== "" && !parts.has(sourcePart))
      unsupported(`Relationship part has no source part: ${part}`);
    for (const relationship of relationships(
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

const escapedMarkdown = (text: string): string =>
  text
    .replaceAll("\\", "\\\\")
    .replaceAll("`", "\\`")
    .replaceAll("*", "\\*")
    .replaceAll("_", "\\_")
    .replaceAll("[", "\\[")
    .replaceAll("]", "\\]")
    .replaceAll("<", "\\<")
    .replaceAll(">", "\\>");

const fromBookmarkName = (name: string | undefined): BlockId | null => {
  if (!name?.startsWith("adx_")) return null;
  const hex = name.slice("adx_".length);
  if (!/^[0-9a-f]{32}$/.test(hex)) return null;
  const candidate = `b_${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20)}`;
  return isBlockId(candidate) ? candidate : null;
};

const isGeneratedBodyBookmark = (name: string): boolean =>
  /^adx_body_\d{6}$/.test(name);

const emittedBookmarkName = (name: string | undefined): string | null =>
  name !== undefined &&
  (fromBookmarkName(name) !== null || isGeneratedBodyBookmark(name))
    ? name
    : null;

const parseParagraphs = (
  xml: string,
  sourcePart: string,
): {
  paragraphs: readonly Paragraph[];
  unsupported: readonly DocxFidelityItem<"unsupported">[];
} => {
  const paragraphs: Paragraph[] = [];
  const unsupportedItems: DocxFidelityItem<"unsupported">[] = [];
  let current: {
    bookmark: BlockId | null;
    bookmarkName: string | null;
    text: string;
    heading: number | null;
    comments: TrackedCommentAnchor[];
  } | null = null;
  const commentStarts: Array<{ id: string; start: number }> = [];
  let runText = "";
  let inText = false;
  let unsupportedDepth = 0;
  const unsupportedKinds = new Set([
    "tbl",
    "ins",
    "del",
    "moveFrom",
    "moveTo",
    "altChunk",
  ]);
  parseDocxXml(
    xml,
    (tag: SaxesTagNS) => {
      if (unsupportedKinds.has(tag.local)) {
        unsupportedDepth++;
        unsupportedItems.push({
          status: "unsupported",
          partPath: sourcePart,
          relationshipId: null,
          ooxmlKind: `w:${tag.local}`,
          count: 1,
          blockIds: [],
          sourcePositions: [],
          explanation:
            "This OOXML construct cannot be represented by the version 1 Markdown importer.",
        });
      }
      if (tag.local === "p")
        current = {
          bookmark: null,
          bookmarkName: null,
          text: "",
          heading: null,
          comments: [],
        };
      if (tag.local === "commentRangeStart" && current) {
        const id = docxXmlAttribute(tag, "id");
        if (!id || !/^\d+$/.test(id) || commentStarts.length > 0)
          unsupported("Comment range nesting is malformed");
        commentStarts.push({ id: id as string, start: current.text.length });
      }
      if (tag.local === "bookmarkStart" && current) {
        const name = emittedBookmarkName(docxXmlAttribute(tag, "name"));
        if (name !== null) {
          if (current.bookmarkName !== null)
            unsupported("Paragraph has multiple agent-docx bookmarks");
          current.bookmarkName = name;
          current.bookmark = fromBookmarkName(name);
        }
      }
      if (tag.local === "pStyle" && current) {
        const style = docxXmlAttribute(tag, "val") ?? "";
        const match = /^Heading([1-6])$/.exec(style);
        if (match) current.heading = Number(match[1]);
      }
      if (tag.local === "r") runText = "";
      if (tag.local === "t" && current) inText = true;
      if (tag.local === "tab" && current) current.text += "\t";
      if ((tag.local === "br" || tag.local === "cr") && current)
        current.text += "\n";
    },
    (tag) => {
      if (tag.local === "t") inText = false;
      if (tag.local === "r" && current) current.text += runText;
      if (tag.local === "commentRangeEnd" && current) {
        const id = docxXmlAttribute(tag, "id");
        const start = commentStarts.pop();
        if (!id || !start || id !== start.id)
          unsupported("Comment range nesting is malformed");
        const safeStart = start as { id: string; start: number };
        current.comments.push({
          id: safeStart.id,
          start: safeStart.start,
          end: current.text.length,
        });
      }
      if (tag.local === "p" && current) {
        if (commentStarts.length > 0)
          unsupported("Comment range crosses a paragraph boundary");
        paragraphs.push({ ...current, sourcePart });
        current = null;
      }
      if (unsupportedKinds.has(tag.local)) unsupportedDepth--;
    },
    (text) => {
      if (inText && current) runText += text;
    },
  );
  if (unsupportedDepth !== 0 || commentStarts.length !== 0)
    throw new AgentDocxError(
      "DOCX_IMPORT_UNSUPPORTED",
      "Malformed OOXML nesting",
    );
  return { paragraphs, unsupported: unsupportedItems };
};

const sourceTextByBlockId = (
  document: LegalDocument,
): ReadonlyMap<BlockId, string> => {
  const text = new Map<BlockId, string>();
  const visit = (blocks: readonly LegalDocument["blocks"][number][]): void => {
    for (const block of blocks) {
      if ("runs" in block)
        text.set(
          block.id,
          block.runs
            .map((run) => `${run.text}${run.hardBreakAfter ? "\n" : ""}`)
            .join(""),
        );
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

/**
 * Reads the non-overlapping tracked-run subset emitted by this package.  The
 * caller projects the paired paragraph text onto source-mapped Markdown; this
 * parser deliberately refuses nested revisions or revisions outside a
 * paragraph rather than flattening an ambiguous Word review stream.
 */
const parseTrackedParagraphs = (
  xml: string,
  sourcePart: string,
): readonly TrackedParagraph[] => {
  const paragraphs: TrackedParagraph[] = [];
  let current: {
    bookmark: BlockId | null;
    baseText: string;
    headText: string;
    visibleText: string;
    heading: number | null;
    revisions: NativeRevision[];
    comments: TrackedCommentAnchor[];
  } | null = null;
  let revision: NativeRevision | null = null;
  const commentStarts: Array<{ id: string; start: number }> = [];
  let inText = false;
  const append = (value: string): void => {
    if (!current) return;
    if (revision) {
      revision.text += value;
      if (revision.kind === "ins" || revision.kind === "moveTo") {
        current.headText += value;
        current.visibleText += value;
      } else current.baseText += value;
      return;
    }
    current.baseText += value;
    current.headText += value;
    current.visibleText += value;
  };
  parseDocxXml(
    xml,
    (tag) => {
      if (tag.local === "p") {
        if (current) unsupported("Nested DOCX paragraphs are unsupported");
        current = {
          bookmark: null,
          baseText: "",
          headText: "",
          visibleText: "",
          heading: null,
          revisions: [],
          comments: [],
        };
      }
      if (!current) {
        if (["ins", "del", "moveFrom", "moveTo"].includes(tag.local))
          unsupported("Tracked revision is outside a paragraph");
        return;
      }
      if (tag.local === "bookmarkStart") {
        const bookmark = fromBookmarkName(docxXmlAttribute(tag, "name"));
        if (bookmark) {
          if (current.bookmark)
            unsupported("Paragraph has multiple block bookmarks");
          current.bookmark = bookmark;
        }
        return;
      }
      if (tag.local === "pStyle") {
        const style = docxXmlAttribute(tag, "val") ?? "";
        const match = /^Heading([1-6])$/.exec(style);
        if (match) current.heading = Number(match[1]);
        return;
      }
      if (tag.local === "commentRangeStart") {
        const id = docxXmlAttribute(tag, "id");
        if (!id || !/^\d+$/.test(id) || commentStarts.length > 0)
          unsupported("Comment range nesting is malformed");
        commentStarts.push({
          id: id as string,
          start: current.visibleText.length,
        });
        return;
      }
      if (["ins", "del", "moveFrom", "moveTo"].includes(tag.local)) {
        if (revision) unsupported("Nested tracked revisions are unsupported");
        const id = docxXmlAttribute(tag, "id");
        const author = docxXmlAttribute(tag, "author");
        const date = docxXmlAttribute(tag, "date");
        if (!id || !/^\d+$/.test(id) || author === undefined)
          unsupported("Tracked revision has invalid native attribution");
        if (date !== undefined && Number.isNaN(new Date(date).valueOf()))
          unsupported("Tracked revision has an invalid native date");
        revision = {
          id: id as string,
          kind: tag.local as NativeRevision["kind"],
          author: author as string,
          date: date ?? null,
          text: "",
        };
        return;
      }
      if (tag.local === "t" || tag.local === "delText") inText = true;
      else if (tag.local === "tab") append("\t");
      else if (tag.local === "br" || tag.local === "cr") append("\n");
    },
    (tag) => {
      if (tag.local === "t" || tag.local === "delText") {
        inText = false;
        return;
      }
      if (tag.local === "commentRangeEnd") {
        const id = docxXmlAttribute(tag, "id");
        const start = commentStarts.pop();
        if (!id || !start || start.id !== id)
          unsupported("Comment range nesting is malformed");
        const safeStart = start as { id: string; start: number };
        current!.comments.push({
          id: safeStart.id,
          start: safeStart.start,
          end: current!.visibleText.length,
        });
        return;
      }
      if (["ins", "del", "moveFrom", "moveTo"].includes(tag.local)) {
        const closedRevision = revision;
        if (!closedRevision || closedRevision.kind !== tag.local)
          unsupported("Tracked revision nesting is malformed");
        current!.revisions.push(closedRevision as NativeRevision);
        revision = null;
        return;
      }
      if (tag.local === "p") {
        if (revision || commentStarts.length > 0)
          unsupported(
            "Tracked revision or comment crosses a paragraph boundary",
          );
        paragraphs.push({ ...current!, sourcePart });
        current = null;
      }
    },
    (text) => {
      if (inText) append(text);
    },
  );
  if (current || revision || inText)
    unsupported("Tracked DOCX text nesting is malformed");
  return paragraphs;
};

const trackedBlockText = (
  block: LegalDocument["blocks"][number] | LegalDocument["footnotes"][number],
): string => {
  if ("runs" in block)
    return block.runs
      .map((run) => `${run.text}${run.hardBreakAfter ? "\n" : ""}`)
      .join("");
  if (block.kind === "footnote")
    return block.paragraphs
      .map((paragraph) =>
        paragraph.runs
          .map((run) => `${run.text}${run.hardBreakAfter ? "\n" : ""}`)
          .join(""),
      )
      .join("\n");
  return "";
};

const reconstructTrackedMaterial = (
  semantic: SemanticManifest,
  paragraphs: readonly TrackedParagraph[],
  assets: Readonly<Record<string, ImportedAsset>>,
): TrackedMaterial => {
  if (semantic.mode !== "redline")
    unsupported("Tracked reconstruction requires a redline semantic manifest");
  const projectedDocument = asObject(
    semantic.document,
    "Semantic manifest document",
  );
  const parsed = parseLegalMarkdown(semantic.source, {
    projectId: semantic.projectId,
    documentId: semantic.documentId,
    metadata: projectedDocument.metadata as LitigationMetadata,
    chrome: (projectedDocument.chrome ?? {}) as DocumentChrome,
    assets,
    requireMarkers: true,
  }).document;
  const byBookmark = new Map<
    string,
    LegalDocument["blocks"][number] | LegalDocument["footnotes"][number]
  >();
  for (const block of [...parsed.blocks, ...parsed.footnotes])
    byBookmark.set(blockBookmark(block.id), block);
  const replacements: Array<{ start: number; end: number; source: string }> =
    [];
  for (const paragraph of paragraphs) {
    if (paragraph.revisions.length === 0) continue;
    const bookmark = paragraph.bookmark;
    if (!bookmark)
      unsupported("Tracked revision has no agent-docx block bookmark");
    const block = byBookmark.get(blockBookmark(bookmark as BlockId));
    if (!block)
      unsupported(
        "Tracked revision bookmark is not declared by the semantic source",
      );
    const resolvedBlock = block as
      | LegalDocument["blocks"][number]
      | LegalDocument["footnotes"][number];
    if (trackedBlockText(resolvedBlock) !== paragraph.headText)
      unsupported("Tracked revision head text does not match semantic source");
    if (paragraph.baseText === paragraph.headText)
      unsupported("Tracked revision has no visible base-to-head change");
    const occurrences =
      resolvedBlock.sourceText.split(paragraph.headText).length - 1;
    if (occurrences !== 1)
      unsupported(
        "Tracked revision cannot be mapped unambiguously to a Markdown source range",
      );
    replacements.push({
      start: resolvedBlock.position.start.offset,
      end: resolvedBlock.position.end.offset,
      source: resolvedBlock.sourceText.replace(
        paragraph.headText,
        paragraph.baseText,
      ),
    });
  }
  if (replacements.length === 0) {
    if (semantic.revisionMap.length > 0)
      unsupported(
        "Redline semantic manifest declares unrepresented tracked revisions",
      );
    return {
      baseSource: semantic.source,
      headSource: semantic.source,
      paragraphs,
    };
  }
  replacements.sort((left, right) => right.start - left.start);
  for (let index = 1; index < replacements.length; index++)
    if (replacements[index - 1]!.start < replacements[index]!.end)
      unsupported("Tracked revisions overlap in the semantic source");
  let baseSource = semantic.source;
  for (const replacement of replacements)
    baseSource =
      baseSource.slice(0, replacement.start) +
      replacement.source +
      baseSource.slice(replacement.end);
  return {
    baseSource,
    headSource: semantic.source,
    paragraphs,
  };
};

const isCodePointBoundary = (value: string, offset: number): boolean =>
  offset >= 0 &&
  offset <= value.length &&
  (offset === 0 ||
    offset === value.length ||
    !(
      value.charCodeAt(offset - 1) >= 0xd800 &&
      value.charCodeAt(offset - 1) <= 0xdbff &&
      value.charCodeAt(offset) >= 0xdc00 &&
      value.charCodeAt(offset) <= 0xdfff
    ));

const trackedAnnotations = (
  semantic: SemanticManifest,
  tracked: TrackedMaterial,
  document: LegalDocument,
  nativeComments: ReadonlyMap<string, NativeComment>,
): readonly ReviewAnnotation[] => {
  const anchors = tracked.paragraphs.flatMap((paragraph) =>
    paragraph.comments.map((anchor) => ({ paragraph, anchor })),
  );
  if (anchors.length !== nativeComments.size)
    unsupported("Native comment definitions and anchors do not match");
  const anchorsById = new Map<string, (typeof anchors)[number]>();
  for (const entry of anchors) {
    if (anchorsById.has(entry.anchor.id))
      unsupported("Native comment has multiple or duplicate anchors");
    anchorsById.set(entry.anchor.id, entry);
  }
  const nativeIds = [...nativeComments.keys()].sort(
    (left, right) => Number(left) - Number(right),
  );
  if (semantic.commentMap.length !== nativeIds.length)
    unsupported("Semantic comment map does not match native comment count");
  const blocks = new Map(
    [...document.blocks, ...document.footnotes].map((block) => [
      block.id,
      block,
    ]),
  );
  return nativeIds.map((nativeId, index) => {
    const native = nativeComments.get(nativeId)!;
    const mapped = semantic.commentMap[index]!;
    const candidate = anchorsById.get(nativeId);
    if (!candidate || !candidate.paragraph.bookmark)
      unsupported("Native comment does not anchor to an agent-docx block");
    const anchor = candidate as (typeof anchors)[number];
    const bookmark = anchor.paragraph.bookmark as BlockId;
    const block = blocks.get(bookmark);
    if (!block) unsupported("Native comment anchors an unknown block");
    const resolvedBlock = block as
      | LegalDocument["blocks"][number]
      | LegalDocument["footnotes"][number];
    const visible = trackedBlockText(resolvedBlock);
    if (
      anchor.anchor.end > visible.length ||
      anchor.anchor.start > anchor.anchor.end ||
      !isCodePointBoundary(visible, anchor.anchor.start) ||
      !isCodePointBoundary(visible, anchor.anchor.end)
    )
      unsupported("Native comment range is not code-point safe");
    const blockWide =
      anchor.anchor.start === 0 && anchor.anchor.end === visible.length;
    if (mapped.blockWide !== blockWide)
      unsupported("Semantic comment map does not match native comment anchor");
    if (!/^a_[0-9a-f-]{36}$/.test(mapped.annotationId))
      unsupported("Semantic comment map has an invalid annotation ID");
    return {
      id: mapped.annotationId as `a_${string}`,
      blockId: resolvedBlock.id,
      ...(blockWide
        ? {}
        : { range: { start: anchor.anchor.start, end: anchor.anchor.end } }),
      author:
        native.author === ""
          ? null
          : {
              name: native.author,
              ...(mapped.authorEmail === null
                ? {}
                : { email: mapped.authorEmail }),
            },
      createdAt: native.date,
      message: native.text,
      status: "open" as const,
    };
  });
};
const redlineAnnotations = (
  semantic: SemanticManifest,
  tracked: TrackedMaterial,
  document: LegalDocument,
  nativeComments: ReadonlyMap<string, NativeComment>,
  visibleTextByBlock: ReadonlyMap<BlockId, string>,
): readonly ReviewAnnotation[] => {
  const anchors = tracked.paragraphs.flatMap((paragraph) =>
    paragraph.comments.map((anchor) => ({ paragraph, anchor })),
  );
  if (anchors.length !== nativeComments.size)
    unsupported("Native comment definitions and anchors do not match");
  const anchorsById = new Map<string, (typeof anchors)[number]>();
  for (const entry of anchors) {
    if (anchorsById.has(entry.anchor.id))
      unsupported("Native comment has multiple or duplicate anchors");
    anchorsById.set(entry.anchor.id, entry);
  }
  const nativeIds = [...nativeComments.keys()].sort(
    (left, right) => Number(left) - Number(right),
  );
  if (nativeIds.length < semantic.commentMap.length)
    unsupported("Semantic comment map has missing native comments");
  const declaredIds = nativeIds.slice(0, semantic.commentMap.length);
  const highestDeclaredId =
    declaredIds.length === 0
      ? -1
      : Math.max(...declaredIds.map((id) => Number(id)));
  if (
    nativeIds
      .slice(semantic.commentMap.length)
      .some((id) => Number(id) <= highestDeclaredId)
  )
    unsupported(
      "Additional reviewer comments cannot be distinguished from declared comments",
    );
  const blocks = new Map(
    [...document.blocks, ...document.footnotes].map((block) => [
      block.id,
      block,
    ]),
  );
  const usedAnnotationIds = new Set(
    semantic.commentMap.map((entry) => entry.annotationId),
  );
  return nativeIds.map((nativeId, index) => {
    const native = nativeComments.get(nativeId)!;
    const candidate = anchorsById.get(nativeId);
    if (!candidate || !candidate.paragraph.bookmark)
      unsupported("Native comment does not anchor to an agent-docx block");
    const anchor = candidate as (typeof anchors)[number];
    const bookmark = anchor.paragraph.bookmark as BlockId;
    const block = blocks.get(bookmark);
    if (!block) unsupported("Native comment anchors an unknown block");
    const resolvedBlock = block as
      | LegalDocument["blocks"][number]
      | LegalDocument["footnotes"][number];
    const visible =
      visibleTextByBlock.get(bookmark) ?? trackedBlockText(resolvedBlock);
    if (
      anchor.anchor.end > visible.length ||
      anchor.anchor.start > anchor.anchor.end ||
      !isCodePointBoundary(visible, anchor.anchor.start) ||
      !isCodePointBoundary(visible, anchor.anchor.end)
    )
      unsupported("Native comment range is not code-point safe");
    const blockWide =
      anchor.anchor.start === 0 && anchor.anchor.end === visible.length;
    const mapped = semantic.commentMap[index];
    if (mapped) {
      if (mapped.blockWide !== blockWide)
        unsupported(
          "Semantic comment map does not match native comment anchor",
        );
      if (!/^a_[0-9a-f-]{36}$/.test(mapped.annotationId))
        unsupported("Semantic comment map has an invalid annotation ID");
      return {
        id: mapped.annotationId as `a_${string}`,
        blockId: resolvedBlock.id,
        ...(blockWide
          ? {}
          : { range: { start: anchor.anchor.start, end: anchor.anchor.end } }),
        author:
          native.author === ""
            ? null
            : {
                name: native.author,
                ...(mapped.authorEmail === null
                  ? {}
                  : { email: mapped.authorEmail }),
              },
        createdAt: native.date,
        message: native.text,
        status: "open" as const,
      };
    }
    const digest = createHash("sha256")
      .update(
        `agent-docx:review-comment:${semantic.projectId}:${semantic.documentId}:${bookmark}:${nativeId}`,
      )
      .digest("hex");
    const annotationId = `a_${digest.slice(0, 8)}-${digest.slice(
      8,
      12,
    )}-${digest.slice(12, 16)}-${digest.slice(16, 20)}-${digest.slice(
      20,
      32,
    )}` as `a_${string}`;
    if (usedAnnotationIds.has(annotationId))
      unsupported("Additional reviewer comment has a duplicate annotation ID");
    usedAnnotationIds.add(annotationId);
    return {
      id: annotationId,
      blockId: resolvedBlock.id,
      ...(blockWide
        ? {}
        : { range: { start: anchor.anchor.start, end: anchor.anchor.end } }),
      author: native.author === "" ? null : { name: native.author },
      createdAt: native.date,
      message: native.text,
      status: "open" as const,
    };
  });
};

type RedlineDecision = "accept" | "reject";
type RedlineDecisionMap = Readonly<Record<string, RedlineDecision>>;

const extendedRevisionEntry = (
  entry: SemanticRevisionMapEntry,
): entry is SemanticRevisionMapEntry & {
  blockId: BlockId | null;
  baseText: string;
  headText: string;
} =>
  entry.blockId !== undefined &&
  entry.baseText !== undefined &&
  entry.headText !== undefined;

const semanticBlocks = (
  document: LegalDocument,
): ReadonlyMap<
  BlockId,
  LegalDocument["blocks"][number] | LegalDocument["footnotes"][number]
> => {
  const blocks = new Map<
    BlockId,
    LegalDocument["blocks"][number] | LegalDocument["footnotes"][number]
  >();
  const visit = (entries: readonly LegalDocument["blocks"][number][]): void => {
    for (const block of entries) {
      blocks.set(block.id, block);
      if (block.kind === "exhibit" || block.kind === "length-exclusion")
        visit(block.blocks);
    }
  };
  visit(document.blocks);
  for (const footnote of document.footnotes) blocks.set(footnote.id, footnote);
  return blocks;
};

const replaceVisibleBlockText = (
  block: LegalDocument["blocks"][number] | LegalDocument["footnotes"][number],
  visibleText: string,
): { start: number; end: number; replacement: string } | null => {
  const current = trackedBlockText(block);
  if (current === visibleText) return null;
  const occurrences = block.sourceText.split(current).length - 1;
  if (current.length === 0 || occurrences !== 1)
    unsupported(
      "Resolved redline paragraph cannot be mapped unambiguously to its Markdown source range",
    );
  return {
    start: block.position.start.offset,
    end: block.position.end.offset,
    replacement: block.sourceText.replace(current, visibleText),
  };
};

const sourceWithVisibleBlocks = (
  source: string,
  blocks: ReadonlyMap<
    BlockId,
    LegalDocument["blocks"][number] | LegalDocument["footnotes"][number]
  >,
  visibleByBlock: ReadonlyMap<BlockId, string>,
): string => {
  const replacements = [...visibleByBlock.entries()].flatMap(
    ([blockId, text]) => {
      const block = blocks.get(blockId);
      if (!block)
        unsupported("Resolved redline paragraph references an unknown block");
      const replacement = replaceVisibleBlockText(
        block as
          | LegalDocument["blocks"][number]
          | LegalDocument["footnotes"][number],
        text,
      );
      return replacement ? [replacement] : [];
    },
  );
  replacements.sort((left, right) => right.start - left.start);
  for (let index = 1; index < replacements.length; index++)
    if (replacements[index - 1]!.start < replacements[index]!.end)
      unsupported("Resolved redline paragraph source ranges overlap");
  let result = source;
  for (const replacement of replacements)
    result =
      result.slice(0, replacement.start) +
      replacement.replacement +
      result.slice(replacement.end);
  return result;
};

const applyRevisionGroup = (
  headText: string,
  entries: readonly (SemanticRevisionMapEntry & {
    blockId: BlockId;
    baseText: string;
    headText: string;
  })[],
  mask: number,
): string | null => {
  const replacements: Array<{
    start: number;
    end: number;
    replacement: string;
  }> = [];
  for (const [index, entry] of entries.entries()) {
    if ((mask & (1 << index)) === 0) continue;
    if (entry.headText.length === 0) return null;
    const positions: number[] = [];
    let cursor = 0;
    while (true) {
      const found = headText.indexOf(entry.headText, cursor);
      if (found < 0) break;
      positions.push(found);
      cursor = found + entry.headText.length;
    }
    if (positions.length !== 1) return null;
    const start = positions[0]!;
    replacements.push({
      start,
      end: start + entry.headText.length,
      replacement: entry.baseText,
    });
  }
  replacements.sort((left, right) => right.start - left.start);
  for (let index = 1; index < replacements.length; index++)
    if (replacements[index - 1]!.start < replacements[index]!.end) return null;
  let result = headText;
  for (const replacement of replacements)
    result =
      result.slice(0, replacement.start) +
      replacement.replacement +
      result.slice(replacement.end);
  return result;
};

const resolveRevisionGroups = (
  semantic: SemanticManifest,
  document: LegalDocument,
  tracked: TrackedMaterial,
): {
  decisions: RedlineDecisionMap;
  visibleByBlock: ReadonlyMap<BlockId, string>;
  baseByBlock: ReadonlyMap<BlockId, string>;
} => {
  const blocks = semanticBlocks(document);
  const byBookmark = new Map<BlockId, TrackedParagraph>();
  for (const paragraph of tracked.paragraphs) {
    if (!paragraph.bookmark) {
      if (paragraph.visibleText.length > 0)
        unsupported("Resolved redline text has no agent-docx block bookmark");
      continue;
    }
    if (byBookmark.has(paragraph.bookmark))
      unsupported("Resolved redline has duplicate agent-docx block bookmarks");
    byBookmark.set(paragraph.bookmark, paragraph);
  }
  const allExtended = semantic.revisionMap.every(extendedRevisionEntry);
  if (!allExtended) {
    for (const [blockId, paragraph] of byBookmark) {
      const block = semanticBlocks(document).get(blockId);
      if (!block)
        unsupported("Resolved redline paragraph references an unknown block");
      if (
        paragraph.visibleText !==
        trackedBlockText(
          block as
            | LegalDocument["blocks"][number]
            | LegalDocument["footnotes"][number],
        )
      )
        unsupported(
          "Resolved redline changes without base/head text are ambiguous",
        );
    }
    return {
      decisions: Object.fromEntries(
        semantic.revisionMap.map((entry) => [entry.changeId, "accept"]),
      ) as RedlineDecisionMap,
      visibleByBlock: new Map(
        [...byBookmark].map(([blockId, paragraph]) => [
          blockId,
          paragraph.visibleText,
        ]),
      ),
      baseByBlock: new Map(),
    };
  }
  const groups = new Map<
    BlockId,
    (SemanticRevisionMapEntry & {
      blockId: BlockId;
      baseText: string;
      headText: string;
    })[]
  >();
  const decisionsSet = new Map<string, RedlineDecision>();
  for (const entry of semantic.revisionMap) {
    if (entry.blockId === null)
      unsupported(
        "Resolved redline change cannot be attributed to an agent-docx block",
      );
    const blockId = entry.blockId as BlockId;
    if (!blocks.has(blockId) || !byBookmark.has(blockId))
      unsupported("Resolved redline change references an unknown block");
    const entries = groups.get(blockId) ?? [];
    entries.push({
      ...entry,
      blockId,
      baseText: entry.baseText as string,
      headText: entry.headText as string,
    });
    groups.set(blockId, entries);
  }
  const visibleByBlock = new Map<BlockId, string>();
  for (const [blockId, paragraph] of byBookmark)
    visibleByBlock.set(blockId, paragraph.visibleText);
  for (const [blockId, paragraph] of byBookmark) {
    const block = blocks.get(blockId);
    if (!block)
      unsupported("Resolved redline paragraph references an unknown block");
    const expectedHead = trackedBlockText(
      block as
        | LegalDocument["blocks"][number]
        | LegalDocument["footnotes"][number],
    );
    const entries = groups.get(blockId) ?? [];
    if (entries.length === 0 && paragraph.visibleText !== expectedHead)
      unsupported(
        "Resolved redline contains a foreign edit outside declared changes",
      );
    if (entries.length === 0) continue;
    if (paragraph.visibleText === expectedHead) {
      for (const entry of entries) decisionsSet.set(entry.changeId, "accept");
      continue;
    }
    if (!allExtended)
      unsupported(
        "Resolved redline changes without base/head text are ambiguous",
      );
    if (entries.length > 20)
      unsupported(
        "Resolved redline paragraph has too many changes to attribute safely",
      );
    const matches: Array<{ mask: number; text: string }> = [];
    const combinations = 1 << entries.length;
    for (let mask = 0; mask < combinations; mask++) {
      const candidate = applyRevisionGroup(expectedHead, entries, mask);
      if (candidate === paragraph.visibleText)
        matches.push({ mask, text: candidate });
    }
    if (matches.length === 0 && entries.length === 1) {
      const entry = entries[0]!;
      if (
        entry.headText.length === 0 &&
        paragraph.visibleText === entry.baseText
      )
        matches.push({ mask: 1, text: paragraph.visibleText });
    }
    if (matches.length !== 1)
      unsupported(
        matches.length === 0
          ? "Resolved redline paragraph does not match any declared accept/reject decision"
          : "Resolved redline paragraph matches multiple accept/reject decisions",
      );
    const mask = matches[0]!.mask;
    for (const [index, entry] of entries.entries())
      decisionsSet.set(
        entry.changeId,
        (mask & (1 << index)) === 0 ? "accept" : "reject",
      );
  }
  const baseByBlock = new Map<BlockId, string>();
  for (const [blockId, entries] of groups) {
    const block = blocks.get(blockId)!;
    const head = trackedBlockText(
      block as
        | LegalDocument["blocks"][number]
        | LegalDocument["footnotes"][number],
    );
    const candidate =
      allExtended && entries.length <= 20
        ? applyRevisionGroup(head, entries, (1 << entries.length) - 1)
        : null;
    if (candidate !== null) baseByBlock.set(blockId, candidate);
  }
  return {
    decisions: Object.fromEntries(decisionsSet) as RedlineDecisionMap,
    visibleByBlock,
    baseByBlock,
  };
};

const cleanAnnotations = (
  paragraphs: readonly Paragraph[],
  document: LegalDocument,
  nativeComments: ReadonlyMap<string, NativeComment>,
): readonly ReviewAnnotation[] => {
  const anchors = paragraphs.flatMap((paragraph) =>
    paragraph.comments.map((anchor) => ({ paragraph, anchor })),
  );
  if (anchors.length !== nativeComments.size)
    unsupported("Native comment definitions and anchors do not match");
  const byNativeId = new Map<string, (typeof anchors)[number]>();
  for (const anchor of anchors) {
    if (byNativeId.has(anchor.anchor.id))
      unsupported("Native comment has multiple or duplicate anchors");
    byNativeId.set(anchor.anchor.id, anchor);
  }
  const blocks = new Map(
    [...document.blocks, ...document.footnotes].map((block) => [
      block.id,
      block,
    ]),
  );
  return [...nativeComments.keys()]
    .sort((left, right) => Number(left) - Number(right))
    .map((nativeId) => {
      const candidate = byNativeId.get(nativeId);
      if (!candidate || !candidate.paragraph.bookmark)
        unsupported("Native comment does not anchor to an agent-docx block");
      const safeCandidate = candidate as (typeof anchors)[number];
      const block = blocks.get(safeCandidate.paragraph.bookmark as BlockId);
      if (!block) unsupported("Native comment anchors an unknown block");
      const safeBlock = block as
        | LegalDocument["blocks"][number]
        | LegalDocument["footnotes"][number];
      const visible = trackedBlockText(safeBlock);
      if (
        safeCandidate.paragraph.text !== visible ||
        safeCandidate.anchor.start > safeCandidate.anchor.end ||
        safeCandidate.anchor.end > visible.length ||
        !isCodePointBoundary(visible, safeCandidate.anchor.start) ||
        !isCodePointBoundary(visible, safeCandidate.anchor.end)
      )
        unsupported("Native comment range is not code-point safe");
      const digest = createHash("sha256")
        .update(`${safeCandidate.paragraph.bookmark}:${nativeId}`)
        .digest("hex");
      const annotationId =
        `a_${digest.slice(0, 8)}-${digest.slice(8, 12)}-${digest.slice(
          12,
          16,
        )}-${digest.slice(16, 20)}-${digest.slice(20, 32)}` as `a_${string}`;
      const native = nativeComments.get(nativeId)!;
      const blockWide =
        safeCandidate.anchor.start === 0 &&
        safeCandidate.anchor.end === visible.length;
      return {
        id: annotationId,
        blockId: safeBlock.id,
        ...(blockWide
          ? {}
          : {
              start: safeCandidate.anchor.start,
              end: safeCandidate.anchor.end,
            }),
        author: native.author === "" ? null : { name: native.author },
        createdAt: native.date,
        message: native.text,
        status: "open" as const,
      };
    });
};

const requirePackage = (
  parts: ReadonlyMap<string, Uint8Array>,
): {
  mainPart: string;
  mainXml: string;
  commentsPart: string | null;
  fidelity: DocxFidelityItem<"unsupported">[];
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
  const root = relationships(decodeDocxXml(rootRels as Uint8Array), "");
  const office = root.filter((entry) => /\/officeDocument$/.test(entry.type));
  if (office.length !== 1 || office[0]!.external)
    unsupported("DOCX has no unique internal main document");
  const mainPart = resolveOpcTarget("", office[0]!.target);
  const main = parts.get(mainPart);
  if (!main) unsupported("Main document relationship is dangling");
  const rels = parts.get(relationshipPart(mainPart));
  const mainRelationships = rels
    ? relationships(decodeDocxXml(rels), mainPart)
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
    fidelity: [],
    semantic,
  };
};

const parseNativeComments = (
  xml: string | undefined,
): ReadonlyMap<string, NativeComment> => {
  const comments = new Map<string, NativeComment>();
  if (!xml) return comments;
  let current: NativeComment | null = null;
  let inText = false;
  let paragraphCount = 0;
  parseDocxXml(
    xml,
    (tag) => {
      if (tag.local === "comment") {
        if (current) unsupported("Nested DOCX comments are unsupported");
        const id = docxXmlAttribute(tag, "id");
        const author = docxXmlAttribute(tag, "author");
        const date = docxXmlAttribute(tag, "date");
        if (
          !id ||
          !/^\d+$/.test(id) ||
          author === undefined ||
          comments.has(id)
        )
          unsupported("DOCX comment has invalid native attribution");
        if (date !== undefined && Number.isNaN(new Date(date).valueOf()))
          unsupported("DOCX comment has an invalid native date");
        current = {
          id: id as string,
          author: author as string,
          date: date ?? null,
          text: "",
        };
        paragraphCount = 0;
        return;
      }
      if (!current) return;
      if (tag.local === "p") {
        if (paragraphCount > 0) current.text += "\n";
        paragraphCount++;
      } else if (tag.local === "t") inText = true;
    },
    (tag) => {
      if (tag.local === "t") {
        inText = false;
        return;
      }
      if (tag.local === "comment") {
        const closedComment = current;
        if (!closedComment) unsupported("DOCX comment nesting is malformed");
        comments.set(
          (closedComment as NativeComment).id,
          closedComment as NativeComment,
        );
        current = null;
      }
    },
    (text) => {
      if (current && inText) current.text += text;
    },
  );
  if (current || inText) unsupported("DOCX comment text nesting is malformed");
  return comments;
};

const loadInput = async (input: string | Uint8Array): Promise<Uint8Array> => {
  if (typeof input !== "string") return input;
  const entry = await lstat(input).catch(() => null);
  if (!entry || !entry.isFile() || entry.isSymbolicLink())
    throw new AgentDocxError(
      "INPUT_NOT_FOUND",
      `DOCX input is not a regular file: ${input}`,
    );
  return readFile(input);
};

const attachmentInventory = (
  parts: ReadonlyMap<string, Uint8Array>,
): Record<
  string,
  { sha256: `sha256:${string}`; mediaType: string; bytes: number }
> => {
  const assets: Record<
    string,
    { sha256: `sha256:${string}`; mediaType: string; bytes: number }
  > = {};
  for (const [path, bytes] of parts) {
    if (!path.startsWith("word/media/")) continue;
    const name = basename(path);
    const mediaType = path.endsWith(".png")
      ? "image/png"
      : /\.jpe?g$/i.test(path)
        ? "image/jpeg"
        : "application/octet-stream";
    assets[name] = {
      sha256: sha256(bytes),
      mediaType,
      bytes: bytes.byteLength,
    };
  }
  return assets;
};

type ImportedAsset = { bytes: Uint8Array; mediaType: string };

type AttachmentResolution = {
  assets: Readonly<Record<string, ImportedAsset>>;
  inventory: Readonly<
    Record<
      string,
      { sha256: `sha256:${string}`; mediaType: string; bytes: number }
    >
  >;
  complete: boolean;
};

const safePayloadPath = (value: string): boolean =>
  value.startsWith("files/") &&
  !value.includes("\\") &&
  value
    .split("/")
    .every(
      (part, index) =>
        index === 0 || (part !== "" && part !== "." && part !== ".."),
    );

const normalizedAttachmentEntries = (
  manifest: AttachmentManifest,
  label: string,
): readonly AttachmentManifest["entries"][number][] => {
  if (!validAttachmentManifest(manifest))
    unsupported(
      `${label} does not have the version-1 attachment manifest shape`,
    );
  const names = new Set<string>();
  const paths = new Set<string>();
  const entries = [...manifest.entries].map((entry) => {
    if (
      entry.name.length === 0 ||
      entry.name.startsWith("/") ||
      entry.name.includes("\\") ||
      entry.name
        .split("/")
        .some((part) => part === "" || part === "." || part === "..") ||
      !safePayloadPath(entry.payloadPath)
    )
      unsupported(`${label} has an unsafe attachment entry`);
    if (names.has(entry.name) || paths.has(entry.payloadPath))
      unsupported(`${label} has duplicate attachment entries`);
    names.add(entry.name);
    paths.add(entry.payloadPath);
    return entry;
  });
  return entries.sort(
    (left, right) =>
      left.name.localeCompare(right.name) ||
      left.payloadPath.localeCompare(right.payloadPath),
  );
};

const sameAttachmentEntries = (
  left: readonly AttachmentManifest["entries"][number][],
  right: readonly AttachmentManifest["entries"][number][],
): boolean =>
  left.length === right.length &&
  left.every(
    (entry, index) =>
      entry.name === right[index]!.name &&
      entry.mediaType === right[index]!.mediaType &&
      entry.byteLength === right[index]!.byteLength &&
      entry.sha256 === right[index]!.sha256 &&
      entry.payloadPath === right[index]!.payloadPath,
  );

const pathInside = (root: string, path: string, label: string): string => {
  const target = resolve(root, path);
  const contained = relative(root, target);
  if (
    contained.length === 0 ||
    isAbsolute(contained) ||
    contained.split(sep).some((part) => part === "..")
  )
    unsupported(`${label} escapes its attachment bundle`);
  return target;
};

const regularAttachmentFile = async (
  path: string,
  label: string,
): Promise<Uint8Array> => {
  const entry = await lstat(path).catch(() => null);
  if (!entry || !entry.isFile() || entry.isSymbolicLink())
    unsupported(`${label} is not a regular nonsymlink file`);
  return readFile(path);
};

const bundleFiles = async (directory: string): Promise<readonly string[]> => {
  const root = await lstat(directory).catch(() => null);
  if (!root || !root.isDirectory() || root.isSymbolicLink())
    unsupported("Attachment bundle directory is not a real directory");
  const visit = async (path: string, prefix: string): Promise<string[]> => {
    const entries = await readdir(path, { withFileTypes: true });
    const files: string[] = [];
    for (const entry of entries) {
      if (entry.isSymbolicLink())
        unsupported("Attachment bundle contains a symbolic link");
      const relativePath = prefix ? `${prefix}/${entry.name}` : entry.name;
      if (entry.isDirectory()) {
        files.push(...(await visit(resolve(path, entry.name), relativePath)));
      } else if (entry.isFile()) {
        files.push(relativePath);
      } else {
        unsupported("Attachment bundle contains a non-regular entry");
      }
    }
    return files;
  };
  return visit(directory, "");
};

const resolveAttachmentBundle = async (
  expected: AttachmentManifest | null,
  bundle: ImportAttachmentBundle | undefined,
): Promise<AttachmentResolution> => {
  if (!expected) {
    if (bundle) unsupported("DOCX has no external attachment inventory");
    return { assets: {}, inventory: {}, complete: true };
  }
  const expectedEntries = normalizedAttachmentEntries(
    expected,
    "DOCX semantic attachment inventory",
  );
  const inventory = Object.fromEntries(
    expectedEntries.map((entry) => [
      entry.name,
      {
        sha256: entry.sha256,
        mediaType: entry.mediaType,
        bytes: entry.byteLength,
      },
    ]),
  );
  if (!bundle) return { assets: {}, inventory, complete: false };
  let supplied: AttachmentManifest;
  let sourceFiles: Readonly<Record<string, ImportedAsset>>;
  if ("directory" in bundle) {
    const manifestPath = pathInside(
      bundle.directory,
      "manifest.json",
      "Attachment manifest",
    );
    const content = await regularAttachmentFile(
      manifestPath,
      "Attachment manifest",
    );
    let parsed: unknown;
    try {
      parsed = JSON.parse(decodeDocxXml(content));
    } catch {
      unsupported("Attachment manifest is not valid JSON");
    }
    if (!validAttachmentManifest(parsed))
      unsupported("Attachment manifest does not have the version-1 shape");
    supplied = parsed as AttachmentManifest;
    const suppliedEntries = normalizedAttachmentEntries(
      supplied,
      "Attachment manifest",
    );
    if (!sameAttachmentEntries(expectedEntries, suppliedEntries))
      unsupported(
        "Attachment bundle manifest does not exactly match the DOCX inventory",
      );
    const allowed = new Set([
      "manifest.json",
      ...suppliedEntries.map((entry) => entry.payloadPath),
    ]);
    const actual = await bundleFiles(bundle.directory);
    if (
      actual.length !== allowed.size ||
      actual.some((path) => !allowed.has(path))
    )
      unsupported("Attachment bundle has missing or extra payload files");
    const files: Record<string, ImportedAsset> = {};
    for (const entry of suppliedEntries) {
      const bytes = await regularAttachmentFile(
        pathInside(bundle.directory, entry.payloadPath, "Attachment payload"),
        `Attachment payload ${entry.name}`,
      );
      files[entry.name] = { bytes, mediaType: entry.mediaType };
    }
    sourceFiles = files;
  } else {
    supplied = bundle.manifest;
    const suppliedEntries = normalizedAttachmentEntries(
      supplied,
      "Attachment manifest",
    );
    if (!sameAttachmentEntries(expectedEntries, suppliedEntries))
      unsupported(
        "Attachment bundle manifest does not exactly match the DOCX inventory",
      );
    if (
      Object.keys(bundle.files).length !== suppliedEntries.length ||
      suppliedEntries.some((entry) => bundle.files[entry.name] === undefined)
    )
      unsupported("Attachment bundle has missing or extra payload files");
    sourceFiles = bundle.files;
  }
  const assets: Record<string, ImportedAsset> = {};
  for (const entry of expectedEntries) {
    const asset = sourceFiles[entry.name];
    if (
      !asset ||
      typeof asset.mediaType !== "string" ||
      !(asset.bytes instanceof Uint8Array) ||
      asset.mediaType !== entry.mediaType ||
      asset.bytes.byteLength !== entry.byteLength ||
      sha256(asset.bytes) !== entry.sha256
    )
      unsupported(
        `Attachment payload does not match manifest entry: ${entry.name}`,
      );
    assets[entry.name] = asset as ImportedAsset;
  }
  return { assets, inventory, complete: true };
};

const embeddedAssets = (
  parts: ReadonlyMap<string, Uint8Array>,
): readonly ImportedAsset[] =>
  [...parts.entries()]
    .filter(([path]) => path.startsWith("word/media/"))
    .map(([path, bytes]) => ({
      bytes,
      mediaType: path.endsWith(".png")
        ? "image/png"
        : /\.jpe?g$/i.test(path)
          ? "image/jpeg"
          : "application/octet-stream",
    }));

const sourceAssetsForSemanticDocument = (
  semantic: SemanticManifest | null,
  external: AttachmentResolution,
  parts: ReadonlyMap<string, Uint8Array>,
): {
  assets: Readonly<Record<string, ImportedAsset>>;
  unresolved: readonly string[];
} => {
  if (!semantic) return { assets: {}, unresolved: [] };
  const documentAssets = asObject(
    semantic.document.assets ?? {},
    "Semantic manifest document assets",
  );
  const candidates = [
    ...embeddedAssets(parts),
    ...Object.values(external.assets),
  ];
  const assets: Record<string, ImportedAsset> = {};
  const unresolved: string[] = [];
  for (const [name, raw] of Object.entries(documentAssets)) {
    const asset = asObject(raw, `Semantic manifest asset ${name}`);
    exactKeys(
      asset,
      ["sha256", "mediaType", "bytes"],
      `Semantic manifest asset ${name}`,
    );
    if (
      typeof asset.sha256 !== "string" ||
      !/^sha256:[a-f0-9]{64}$/.test(asset.sha256) ||
      typeof asset.mediaType !== "string" ||
      !Number.isSafeInteger(asset.bytes) ||
      (asset.bytes as number) < 0
    )
      unsupported(`Semantic manifest asset ${name} is invalid`);
    const match = candidates.find(
      (candidate) =>
        candidate.mediaType === asset.mediaType &&
        candidate.bytes.byteLength === asset.bytes &&
        sha256(candidate.bytes) === asset.sha256,
    );
    if (match) {
      assets[name] = match;
    } else {
      unresolved.push(name);
      assets[name] = {
        bytes: new Uint8Array(),
        mediaType: asset.mediaType as string,
      };
    }
  }
  return { assets, unresolved };
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
  const bytes = await loadInput(input);
  const parts = await readDocxParts(bytes);
  const { mainPart, mainXml, commentsPart, semantic } = requirePackage(parts);
  const attachments = await resolveAttachmentBundle(
    semantic?.attachments ?? null,
    _options.attachments,
  );
  const sourceAssets = sourceAssetsForSemanticDocument(
    semantic,
    attachments,
    parts,
  );
  const tracked =
    semantic?.mode === "redline"
      ? reconstructTrackedMaterial(
          semantic,
          parseTrackedParagraphs(mainXml, mainPart),
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
          unsupported: [] as readonly DocxFidelityItem<"unsupported">[],
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
            relationshipId: "rIdAgentDocxSemantic",
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
            relationshipId: "rIdAgentDocxSemantic",
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
              relationshipId: "rIdAgentDocxSemantic",
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
      relationshipId: "rIdAgentDocxSemantic",
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
      sourceSha256: unsupported ? null : sha256(source),
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
  const bytes = await loadInput(input);
  const parts = await readDocxParts(bytes);
  const packageInfo = requirePackage(parts);
  if (!packageInfo.semantic)
    unsupported("Redline resolution requires an agent-docx semantic manifest");
  const semantic = packageInfo.semantic as SemanticManifest;
  if (semantic.mode !== "redline")
    unsupported("Redline resolution requires a redline semantic manifest");
  const attachments = await resolveAttachmentBundle(
    semantic.attachments,
    options.attachments,
  );
  const sourceAssets = sourceAssetsForSemanticDocument(
    semantic,
    attachments,
    parts,
  );
  if (sourceAssets.unresolved.length > 0)
    unsupported(
      `Semantic asset cannot be resolved: ${sourceAssets.unresolved.join(", ")}`,
    );
  const paragraphs = parseTrackedParagraphs(
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
  );
  const baseSource =
    resolved.baseByBlock.size === 0
      ? semantic.source
      : sourceWithVisibleBlocks(semantic.source, blocks, resolved.baseByBlock);
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
