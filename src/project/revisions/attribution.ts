import { AgentDocxError } from "../../types.js";
import {
  readObject,
  readRevisionJson,
  strictJson,
  type OpenedStore,
} from "../store.js";
import { assertStoredConfig, assertStoredDocument } from "../documents.js";
import { defaultAttribution } from "../../revisions/diff.js";
import { visibleTextForBlock } from "../../legal/visible-text.js";
import { applyRevisionDelta } from "./deltas.js";
import type {
  AddressableBlock,
  LegalDocument,
  ReviewAnnotation,
  RevisionId,
} from "../../legal/model.js";
import type {
  AttributionSpan,
  Change,
  ChangeAttribution,
  RevisionRecord,
} from "../../revisions/types.js";
import type { AgentDocxDocumentConfig } from "../contracts.js";

export type RevisionMaterial = {
  revision: RevisionRecord;
  source: string;
  config: AgentDocxDocumentConfig;
  document: LegalDocument;
  annotations: readonly ReviewAnnotation[];
};

export type AttributionState = {
  blocks: Map<string, readonly AttributionSpan[]>;
  operations: Map<string, ChangeAttribution>;
  config: Map<string, ChangeAttribution>;
  configOperations: Map<string, ChangeAttribution>;
  dependencies: Map<string, ChangeAttribution>;
  dependencyOperations: Map<string, ChangeAttribution>;
};

export type MutableJsonObject = Record<string, unknown>;

export type RawReplacement = {
  start: number;
  end: number;
  expectedText: string;
  replacement: string;
};

export type SourceMarkerLine = { id: string; start: number; end: number };

export const provenanceForRevision = async (
  opened: OpenedStore,
  target: RevisionRecord,
): Promise<AttributionState> => {
  const chain: RevisionRecord[] = [];
  const visited = new Set<RevisionId>();
  let current: RevisionRecord | null = target;
  while (current) {
    if (current.documentId !== target.documentId)
      throw new AgentDocxError(
        "PROJECT_INVALID",
        `Revision belongs to another document: ${current.id}`,
      );
    if (visited.has(current.id))
      throw new AgentDocxError(
        "PROJECT_INVALID",
        "Revision graph contains a cycle",
      );
    visited.add(current.id);
    chain.push(current);
    const parent: RevisionId | undefined = current.parents[0];
    current = parent
      ? await readRevisionJson<RevisionRecord>(opened.storePath, parent)
      : null;
  }
  chain.reverse();
  const root = chain[0]!;
  const rootMaterial = await (async () => {
    const config = assertStoredConfig(
      strictJson(
        await readObject(opened.storePath, root.documentConfigObject),
        root.documentConfigObject,
      ),
      root.documentConfigObject,
    );
    const document = assertStoredDocument(
      strictJson(
        await readObject(opened.storePath, root.legalDocumentObject),
        root.legalDocumentObject,
      ),
      root.legalDocumentObject,
    );
    return { config, document };
  })();
  const state = seedAttributionState(
    rootMaterial.document,
    rootMaterial.config,
    defaultAttribution(root.author, root.createdAt),
  );
  for (const key of Object.keys(root.dependencyObjects)) {
    const attribution = defaultAttribution(root.author, root.createdAt);
    state.dependencies.set(key, attribution);
    state.dependencyOperations.set(key, attribution);
  }
  let previousDocument = rootMaterial.document;
  for (const [index, record] of chain.slice(1).entries()) {
    const parent = chain[index]!;
    if (
      record.documentId !== target.documentId ||
      record.parents[0] !== parent.id
    )
      throw new AgentDocxError(
        "PROJECT_INVALID",
        `Revision delta parent is invalid: ${record.id}`,
      );
    if (!record.deltaObject)
      throw new AgentDocxError(
        "PROJECT_INVALID",
        `Revision delta is missing: ${record.id}`,
      );
    const delta = strictJson(
      await readObject(opened.storePath, record.deltaObject),
      record.deltaObject,
    ) as {
      schemaVersion?: unknown;
      parentSourceObject?: unknown;
      parentDocumentConfigObject?: unknown;
      changes?: readonly Change[];
      annotations?: readonly unknown[];
    };
    if (
      delta.schemaVersion !== 1 ||
      delta.parentSourceObject !== parent.sourceObject ||
      delta.parentDocumentConfigObject !== parent.documentConfigObject ||
      !Array.isArray(delta.changes) ||
      !Array.isArray(delta.annotations)
    )
      throw new AgentDocxError(
        "PROJECT_INVALID",
        `Revision delta is malformed: ${record.id}`,
      );
    const currentDocument = JSON.parse(
      new TextDecoder("utf-8", { fatal: true }).decode(
        await readObject(opened.storePath, record.legalDocumentObject),
      ),
    ) as LegalDocument;
    applyRevisionDelta(state, delta.changes, previousDocument, currentDocument);
    previousDocument = currentDocument;
  }
  return state;
};

export const provenanceBlocks = (
  document: LegalDocument,
): readonly AddressableBlock[] => {
  const result: AddressableBlock[] = [];
  const visit = (blocks: readonly AddressableBlock[]): void => {
    for (const block of blocks) {
      result.push(block);
      if (block.kind === "exhibit" || block.kind === "length-exclusion")
        visit(block.blocks);
      else if (block.kind === "list")
        for (const item of block.items) visit(item.children);
    }
  };
  visit(document.blocks);
  result.push(...document.footnotes);
  return result;
};

export const textSpans = (
  text: string,
  attribution: ChangeAttribution,
): readonly AttributionSpan[] =>
  text.length === 0 ? [] : [{ start: 0, end: text.length, attribution }];

export const seedAttributionState = (
  document: LegalDocument,
  config: AgentDocxDocumentConfig,
  attribution: ChangeAttribution,
): AttributionState => {
  const blocks = new Map<string, readonly AttributionSpan[]>();
  const operations = new Map<string, ChangeAttribution>();
  for (const block of provenanceBlocks(document)) {
    const spans = textSpans(visibleTextForBlock(block), attribution);
    blocks.set(block.id, spans);
    operations.set(block.id, attribution);
  }
  const configMap = new Map<string, ChangeAttribution>();
  const seedConfigAttribution = (
    value: unknown,
    path: string,
    valueAttribution: ChangeAttribution,
  ): void => {
    configMap.set(path, valueAttribution);
    if (value !== null && typeof value === "object" && !Array.isArray(value)) {
      for (const [key, child] of Object.entries(
        value as Record<string, unknown>,
      ))
        seedConfigAttribution(
          child,
          `${path}/${key.replaceAll("~", "~0").replaceAll("/", "~1")}`,
          valueAttribution,
        );
    }
  };
  seedConfigAttribution(config, "", attribution);
  const configOperations = new Map(
    [...configMap.keys()].map((path) => [path, attribution] as const),
  );
  return {
    blocks,
    operations,
    config: configMap,
    configOperations,
    dependencies: new Map(),
    dependencyOperations: new Map(),
  };
};

export const setConfigAttribution = (
  config: Map<string, ChangeAttribution>,
  value: unknown,
  path: string,
  attribution: ChangeAttribution,
): void => {
  config.set(path, attribution);
  if (value !== null && typeof value === "object" && !Array.isArray(value)) {
    for (const [key, child] of Object.entries(value as Record<string, unknown>))
      setConfigAttribution(
        config,
        child,
        `${path}/${key.replaceAll("~", "~0").replaceAll("/", "~1")}`,
        attribution,
      );
  }
};

export const removeConfigAttribution = (
  config: Map<string, ChangeAttribution>,
  path: string,
): void => {
  for (const key of config.keys())
    if (key === path || key.startsWith(`${path}/`)) config.delete(key);
};
