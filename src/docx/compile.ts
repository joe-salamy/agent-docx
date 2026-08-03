import { createHash } from "node:crypto";
import canonicalize from "canonicalize";
import { lowerLegalDocument } from "../legal/lower.js";
import type {
  LegalBlock,
  LegalDocument,
  LegalDocumentSpecification,
  ReviewAnnotation,
} from "../legal/model.js";
import { blockBookmark } from "../legal/model.js";
import { parseLegalMarkdown } from "../legal/parse.js";
import { validateLegalDocument, type ValidationResult } from "../legal/rules.js";
import { measureNormalizedDocument } from "../renderers/index.js";
import { AgentDocxError, type MeasurementResult } from "../types.js";
import {
  generateDocx,
  type GeneratedDocx,
  type GenerateDocxOptions,
} from "./generate.js";
import { inspectDocxTemplate } from "./inspect.js";
import type {
  ArtifactResult,
  AttachmentManifest,
  BodyBlockManifestEntry,
  CompiledDocx,
  GeneratedAttachmentBundle,
} from "./contracts.js";
import type { CompileOptions } from "../project/contracts.js";
import type { ChangeSet } from "../revisions/types.js";


export type CompileMarkdownOptions = CompileOptions & {
  generation?: Pick<
    GenerateDocxOptions,
    "revision" | "changeSet" | "annotations" | "dependencies" | "createdAt"
  >;
};
const sha256 = (bytes: Uint8Array | string): `sha256:${string}` =>
  `sha256:${createHash("sha256").update(bytes).digest("hex")}`;

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
      ["source", "sourceText", "segments", "position", "preview", "annotations"].includes(
        key,
      )
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
    return [{
      id: block.id,
      bookmark: entry.id,
      index: entry.index,
      parentId: context.parentId,
      depth: context.depth + (block.kind === "blockquote" ? block.depth : 0),
      kind: block.kind,
      position: block.position,
      preview: entry.preview,
    }];
  });
};

const serializableMeasurement = (
  measurement: MeasurementResult,
): Omit<MeasurementResult, "generatedDocx"> => {
  const { generatedDocx: _generatedDocx, ...serializable } = measurement;
  return serializable;
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
  return [...sources].sort((left, right) => left.localeCompare(right));
};

const attachmentBundle = (
  document: LegalDocument,
  assets: Readonly<Record<string, { bytes: Uint8Array; mediaType: string }>> | undefined,
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
      sha256: sha256(asset.bytes),
      payloadPath: `files/${name}`,
    };
  });
  const manifest = { schemaVersion: 1 as const, entries };
  return {
    manifestSha256: sha256(canonicalize(manifest)!),
    manifest,
    files,
  };
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
  sourceSha256: sha256(input.source),
  document: semanticDocumentProjection(input.document),
  blocks: flattenBlocks(input.document.blocks).map(
    ({ block, parentId, depth }, order) => ({
      id: block.id,
      bookmark: blockBookmark(block.id),
      parentId,
      depth,
      order,
      kind: block.kind,
    }),
  ),
  attachments: input.attachments,
  revision: input.revision,
  baseRevision: input.baseRevision,
  validation: input.validation,
  dependencies: input.dependencies
    ? [...input.dependencies]
        .sort(([left], [right]) => left.localeCompare(right))
        .map(([key, dependency]) => ({
          key,
          sha256: dependency.sha256,
          mediaType: dependency.mediaType,
          byteLength: dependency.bytes.byteLength,
        }))
    : [],
  revisionMap: input.changeSet
    ? [...input.changeSet.changes]
        .sort((left, right) => left.id.localeCompare(right.id))
        .map((change) => ({
          changeId: change.id,
          attribution: {
            author: change.attribution.author ?? null,
            createdAt: change.attribution.createdAt,
            ...(change.attribution.sourceRevisionId
              ? { sourceRevisionId: change.attribution.sourceRevisionId }
              : {}),
          },
        }))
    : [],
  commentMap: [...(input.annotations ?? [])]
    .filter((annotation) => annotation.status === "open")
    .sort((left, right) => left.id.localeCompare(right.id))
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
): Promise<CompiledDocx> => {
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
  for (const marker of [...parsed.missingMarkers].sort((left, right) => right.offset - left.offset))
    sourceWithMarkers = `${sourceWithMarkers.slice(0, marker.offset)}<!-- agent-docx:block id="${marker.id}" -->\n${sourceWithMarkers.slice(marker.offset)}`;
  const attachments = attachmentBundle(document, specification.assets);
  const { generation, ...measurementOptions } = options;
  const measurement = await measureNormalizedDocument(lowerLegalDocument(document), {
    ...measurementOptions,
    profile: specification.profile,
    filingKind: specification.filingKind,
    fontSet: specification.fontSet,
    chrome: document.chrome,
    ...(template ? { template } : {}),
  });
  const validation = validateLegalDocument(document, {
    ...(generation?.revision ? { revision: generation.revision.id } : {}),
    rulePack: specification.rulePack,
    filingKind: specification.filingKind,
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
    dependencies: generation?.dependencies,
    changeSet: generation?.changeSet,
    annotations: generation?.annotations,
  });
  const generated = await generateDocx(
    document,
    measurement.deterministic.profile,
    {
      assets: specification.assets,
      chrome: document.chrome,
      pageCount: Math.max(1, measurement.deterministic.pageCount),
      metadata: document.metadata,
      ...generation,
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
    docxSha256: sha256(generated.bytes),
    dependencies: Object.entries(document.assets)
      .sort(([left], [right]) => left.localeCompare(right))
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
    mediaType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    byteLength: generated.bytes.byteLength,
    sha256: sha256(generated.bytes),
    provenanceSha256: sha256(canonicalize(provenance)!),
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
