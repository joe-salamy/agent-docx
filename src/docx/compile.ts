import canonicalize from "canonicalize";
import type {
  AddressableBlock,
  LegalBlock,
  LegalDocument,
  LegalDocumentSpecification,
  ReviewAnnotation,
} from "../legal/model.js";
import { blockBookmark } from "../legal/model.js";
import { parseLegalMarkdown } from "../legal/parse.js";
import {
  validateLegalDocument,
  type ValidationResult,
} from "../legal/rules.js";
import { measureNormalizedDocument } from "../renderers/index.js";
import { loadFonts } from "../resolve.js";
import { AgentDocxError } from "../types.js";
import { serializableMeasurement } from "../measurement.js";
import {
  generateDocx,
  type GeneratedDocx,
  type GenerateDocxOptions,
} from "./generate.js";
import { inspectDocxTemplate } from "./inspect.js";
import { sha256Hex } from "./package.js";
import { lowerLegalDocument } from "../legal/lower.js";
import type {
  ArtifactResult,
  AttachmentManifest,
  BodyBlockManifestEntry,
  GeneratedAttachmentBundle,
  StatelessCompiledDocx,
} from "./contracts.js";
import type { CompileOptions } from "../project/contracts.js";
import type { Change, ChangeSet } from "../revisions/types.js";
import { definedProps } from "../json-contract.js";
import { visibleTextForBlock } from "../legal/visible-text.js";
import { codePointCompare } from "./helpers.js";
export type CompileMarkdownOptions = CompileOptions & {
  generation?: Pick<
    GenerateDocxOptions,
    "revision" | "changeSet" | "annotations" | "dependencies" | "createdAt"
  >;
};
type FlattenedBlock = {
  block: LegalBlock;
  parentId: LegalBlock["id"] | null;
  depth: number;
};

const flattenBlocks = (
  blocks: readonly LegalBlock[],
  parentId: LegalBlock["id"] | null = null,
  depth = 0,
): readonly FlattenedBlock[] => {
  const flattened: FlattenedBlock[] = [];
  for (const block of blocks) {
    flattened.push({ block, parentId, depth });
    if (block.kind === "exhibit" || block.kind === "length-exclusion")
      flattened.push(...flattenBlocks(block.blocks, block.id, depth + 1));
  }
  return flattened;
};

export const semanticDocumentProjection = (value: unknown): unknown => {
  if (Array.isArray(value)) return value.map(semanticDocumentProjection);
  if (value === null || typeof value !== "object") return value;
  const projection: Record<string, unknown> = {};
  for (const [key, child] of Object.entries(value as Record<string, unknown>)) {
    if (
      [
        "source",
        "sourceText",
        "segments",
        "position",
        "preview",
        "annotations",
      ].includes(key)
    )
      continue;
    projection[key] = semanticDocumentProjection(child);
  }
  return projection;
};

const bodyManifest = (
  document: LegalDocument,
  generated: GeneratedDocx,
): readonly BodyBlockManifestEntry[] => {
  const byBookmark = new Map<string, FlattenedBlock>();
  for (const entry of flattenBlocks(document.blocks))
    byBookmark.set(blockBookmark(entry.block.id), entry);
  return generated.bodyParagraphs.flatMap((entry) => {
    const context = byBookmark.get(entry.id);
    if (!context) return [];
    const { block } = context;
    return [
      {
        id: block.id,
        bookmark: entry.id,
        index: entry.index,
        parentId: context.parentId,
        depth: context.depth + (block.kind === "blockquote" ? block.depth : 0),
        kind: block.kind,
        position: block.position,
        preview: entry.preview,
      },
    ];
  });
};

const exhibitSources = (blocks: readonly LegalBlock[]): readonly string[] => {
  const sources = new Set<string>();
  const visit = (items: readonly LegalBlock[]): void => {
    for (const block of items) {
      if (block.kind === "exhibit") {
        sources.add(block.source);
        visit(block.blocks);
      } else if (block.kind === "length-exclusion") {
        visit(block.blocks);
      } else if (block.kind === "list") {
        for (const item of block.items) visit(item.children);
      }
    }
  };
  visit(blocks);
  return [...sources].sort((left, right) => codePointCompare(left, right));
};

const attachmentBundle = (
  document: LegalDocument,
  assets:
    | Readonly<Record<string, { bytes: Uint8Array; mediaType: string }>>
    | undefined,
): GeneratedAttachmentBundle | null => {
  const names = exhibitSources(document.blocks);
  if (names.length === 0) return null;
  const files: Record<string, { bytes: Uint8Array; mediaType: string }> = {};
  const entries = names.map((name) => {
    const asset = assets?.[name];
    if (!asset)
      throw new AgentDocxError(
        "REFERENCE_INVALID",
        `Missing exhibit attachment: ${name}`,
      );
    files[name] = asset;
    return {
      name,
      mediaType: asset.mediaType,
      byteLength: asset.bytes.byteLength,
      sha256: sha256Hex(asset.bytes),
      payloadPath: `files/${name}`,
    };
  });
  const manifest = { schemaVersion: 1 as const, entries };
  return {
    manifestSha256: sha256Hex(canonicalize(manifest)!),
    manifest,
    files,
  };
};

const blockChildrenForParent = (
  document: LegalDocument,
  collection: "body" | "footnotes",
  parentId: string | null,
): readonly LegalBlock[] => {
  const root: readonly AddressableBlock[] =
    collection === "body" ? document.blocks : document.footnotes;
  if (parentId === null) return root as readonly LegalBlock[];
  const visit = (
    items: readonly AddressableBlock[],
  ): readonly LegalBlock[] | null => {
    for (const block of items) {
      if (block.id === parentId) {
        if (block.kind === "exhibit" || block.kind === "length-exclusion")
          return block.blocks;
        if (block.kind === "list")
          return block.items.flatMap((item) => item.children);
        return [];
      }
      const nested =
        block.kind === "exhibit" || block.kind === "length-exclusion"
          ? block.blocks
          : block.kind === "list"
            ? block.items.flatMap((item) => item.children)
            : [];
      const found = visit(nested);
      if (found !== null) return found;
    }
    return null;
  };
  return visit(root) ?? [];
};

const deletedSourceAnchor = (
  document: LegalDocument,
  change: Extract<Change, { kind: "delete-block" }>,
): number => {
  const siblings = blockChildrenForParent(
    document,
    change.from.collection,
    change.from.parentId,
  );
  const next = siblings[change.from.index];
  if (next) return next.position.start.offset;
  const previous = siblings[siblings.length - 1];
  return previous?.position.end.offset ?? change.oldSource.start;
};

const revisionMapText = (
  change: Change,
  document: LegalDocument,
): {
  blockId: string | null;
  baseText: string;
  headText: string;
  sourceStart?: number;
  sourceEnd?: number;
} => {
  switch (change.kind) {
    case "insert-text":
      return {
        blockId: change.blockId,
        baseText: "",
        headText: change.newText,
      };
    case "delete-text":
      return {
        blockId: change.blockId,
        baseText: change.oldText,
        headText: "",
      };
    case "replace-text":
      return {
        blockId: change.blockId,
        baseText: change.oldText,
        headText: change.newText,
      };
    case "insert-block":
      return {
        blockId: change.blockId,
        baseText: "",
        headText: visibleTextForBlock(change.block),
      };
    case "delete-block":
      return {
        blockId: change.blockId,
        baseText: visibleTextForBlock(change.oldBlock),
        headText: "",
        sourceStart: deletedSourceAnchor(document, change),
        sourceEnd: deletedSourceAnchor(document, change),
      };
    case "replace-block":
      return {
        blockId: change.blockId,
        baseText: visibleTextForBlock(change.oldBlock),
        headText: visibleTextForBlock(change.newBlock),
      };
    case "move-block":
      return {
        blockId: change.blockId,
        baseText: change.oldSource.text,
        headText: change.newSource.text,
      };
    case "replace-container-shell":
      return {
        blockId: change.blockId,
        baseText: change.oldShell.sourceRanges
          .map((range) => range.text)
          .join("\n"),
        headText: change.newShell.sourceRanges
          .map((range) => range.text)
          .join("\n"),
      };
    default:
      return { blockId: null, baseText: "", headText: "" };
  }
};

export const createSemanticManifest = (input: {
  document: LegalDocument;
  source: string;
  mode: "clean" | "redline";
  attachments: AttachmentManifest | null;
  revision: `sha256:${string}` | null;
  baseRevision: `sha256:${string}` | null;
  validation: ValidationResult;
  dependencies?: GenerateDocxOptions["dependencies"];
  changeSet?: ChangeSet;
  annotations?: readonly ReviewAnnotation[];
}): Record<string, unknown> => ({
  schemaVersion: 1,
  generator: "agent-docx",
  mode: input.mode,
  projectId: input.document.projectId,
  documentId: input.document.documentId,
  source: input.source,
  sourceSha256: sha256Hex(input.source),
  document: semanticDocumentProjection(input.document),
  blocks: flattenBlocks(input.document.blocks).map(
    ({ block, parentId, depth }, order) => {
      const authorities =
        "runs" in block
          ? block.runs.flatMap((run, runIndex) =>
              run.authority
                ? [
                    {
                      run: runIndex,
                      id: run.authority.id,
                      category: run.authority.category,
                      short: run.authority.short,
                    },
                  ]
                : [],
            )
          : [];
      return {
        id: block.id,
        bookmark: blockBookmark(block.id),
        parentId,
        depth,
        order,
        kind: block.kind,
        ...(authorities.length > 0 ? { authorities } : {}),
      };
    },
  ),
  attachments: input.attachments,
  revision: input.revision,
  baseRevision: input.baseRevision,
  validation: input.validation,
  dependencies: input.dependencies
    ? [...input.dependencies]
        .sort(([left], [right]) => codePointCompare(left, right))
        .map(([key, dependency]) => ({
          key,
          sha256: dependency.sha256,
          mediaType: dependency.mediaType,
          byteLength: dependency.bytes.byteLength,
        }))
    : [],
  revisionMap: input.changeSet
    ? [...input.changeSet.changes]
        .sort((left, right) => codePointCompare(left.id, right.id))
        .map((change) => {
          const text = revisionMapText(change, input.document);
          return {
            changeId: change.id,
            attribution: {
              author: change.attribution.author ?? null,
              createdAt: change.attribution.createdAt,
              ...(change.attribution.sourceRevisionId
                ? { sourceRevisionId: change.attribution.sourceRevisionId }
                : {}),
            },
            ...text,
          };
        })
    : [],
  commentMap: [...(input.annotations ?? [])]
    .filter((annotation) => annotation.status === "open")
    .sort((left, right) => codePointCompare(left.id, right.id))
    .map((annotation) => ({
      annotationId: annotation.id,
      blockWide: annotation.range === undefined,
      authorEmail: annotation.author?.email ?? null,
    })),
});

export const compileMarkdown = async (
  markdown: string,
  specification: LegalDocumentSpecification,
  options: CompileMarkdownOptions = {},
): Promise<StatelessCompiledDocx> => {
  const template = specification.template
    ? await inspectDocxTemplate(specification.template, {
        fallbackProfile:
          typeof specification.profile === "string"
            ? specification.profile
            : "us-district-conventional",
      })
    : undefined;
  const parsed = parseLegalMarkdown(markdown, specification);
  const document = parsed.document;
  let sourceWithMarkers = markdown;
  for (const marker of [...parsed.missingMarkers].sort(
    (left, right) => right.offset - left.offset,
  ))
    sourceWithMarkers = `${sourceWithMarkers.slice(0, marker.offset)}<!-- agent-docx:block id="${marker.id}" -->\n${sourceWithMarkers.slice(marker.offset)}`;
  const attachments = attachmentBundle(document, specification.assets);
  const { generation, ...measurementOptions } = options;
  const measurement = await measureNormalizedDocument(
    lowerLegalDocument(document),
    {
      ...definedProps(measurementOptions),
      profile: specification.profile,
      ...(specification.filingKind !== undefined
        ? { filingKind: specification.filingKind }
        : {}),
      ...(specification.fontSet !== undefined
        ? { fontSet: specification.fontSet }
        : {}),
      ...(document.chrome !== undefined ? { chrome: document.chrome } : {}),
      ...(template ? { template } : {}),
    },
  );
  const fonts = await loadFonts(
    specification.fontSet,
    measurement.deterministic.profile.requestedFontFamily,
  );
  const validation = validateLegalDocument(document, {
    ...(generation?.revision ? { revision: generation.revision.id } : {}),
    ...(specification.rulePack !== undefined
      ? { rulePack: specification.rulePack }
      : {}),
    ...(options.rulePacks !== undefined
      ? { customPacks: options.rulePacks }
      : {}),
    ...(specification.filingKind !== undefined
      ? { filingKind: specification.filingKind }
      : {}),
    measurement: serializableMeasurement(measurement),
  });
  const semanticManifest = createSemanticManifest({
    document,
    source: sourceWithMarkers,
    mode: "clean",
    attachments: attachments?.manifest ?? null,
    revision: generation?.revision?.id ?? null,
    baseRevision: generation?.changeSet?.baseRevision ?? null,
    validation,
    ...(generation?.dependencies
      ? { dependencies: generation.dependencies }
      : {}),
    ...(generation?.changeSet ? { changeSet: generation.changeSet } : {}),
    ...(generation?.annotations ? { annotations: generation.annotations } : {}),
  });
  const generated = await generateDocx(
    document,
    measurement.deterministic.profile,
    {
      ...(specification.assets !== undefined
        ? { assets: specification.assets }
        : {}),
      fonts,
      pageCount: Math.max(1, measurement.deterministic.pageCount),
      metadata: document.metadata,
      ...(generation ? definedProps(generation) : {}),
      validation,
      semanticManifest,
    },
  );
  const provenance = {
    generator: "agent-docx",
    documentId: specification.documentId,
    mode: "clean",
    profile: measurement.deterministic.profile.id,
    rulePack: specification.rulePack ?? null,
    docxSha256: sha256Hex(generated.bytes),
    dependencies: Object.entries(document.assets)
      .sort(([left], [right]) => codePointCompare(left, right))
      .map(([name, asset]) => ({ name, sha256: asset.sha256 })),
    attachmentManifestSha256: attachments?.manifestSha256 ?? null,
  };
  const artifactAttachments: Extract<
    ArtifactResult,
    { path: null }
  >["attachments"] = attachments
    ? {
        path: null,
        storePath: null,
        manifestSha256: attachments.manifestSha256,
        manifest: attachments.manifest,
      }
    : null;
  const artifact: Extract<ArtifactResult, { path: null }> = {
    schemaVersion: 1,
    mediaType:
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    byteLength: generated.bytes.byteLength,
    sha256: sha256Hex(generated.bytes),
    provenanceSha256: sha256Hex(canonicalize(provenance)!),
    documentId: specification.documentId,
    profile: measurement.deterministic.profile.id,
    rulePack: specification.rulePack ?? null,
    rendererProvenance: {
      generator: "agent-docx",
      requested: options.renderer ?? "deterministic",
      pageCountSource: measurement.pageCountSource,
    },
    path: null,
    storePath: null,
    attachments: artifactAttachments,
    revision: null,
    mode: "clean",
    baseRevision: null,
  };
  return {
    schemaVersion: 1,
    bytes: generated.bytes,
    attachments,
    validation,
    blocks: bodyManifest(document, generated),
    measurement: serializableMeasurement(measurement),
    artifact,
  };
};
