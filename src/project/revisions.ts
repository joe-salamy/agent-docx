import { AgentDocxError } from "../types.js";
import {
  canonicalJson,
  objectId,
  readHead,
  readObject,
  readProjectFile,
  readRevisionJson,
  removeOwnedFile,
  replaceOwnedFile,
  snapshotProjectDocument,
  strictJson,
  withLockedStore,
  type OpenedStore,
  type ProjectSnapshot,
} from "./store.js";
import {
  assertStoredAnnotations,
  assertStoredConfig,
  assertStoredDocument,
  commitLocked,
  documentFor,
} from "./documents.js";
import {
  createChangeSet,
  defaultAttribution,
  reattributeChangeSet,
  reattributeVisibleText,
  type ChangeSetProvenance,
  type JsonObject,
} from "../revisions/diff.js";
import { visibleTextForBlock } from "../legal/visible-text.js";
import { lstat, mkdir } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { pathsOverlap } from "./fs-util.js";
import type {
  Actor,
  AddressableBlock,
  LegalDocument,
  ReviewAnnotation,
  RevisionId,
} from "../legal/model.js";
import type {
  AttributionSpan,
  Change,
  ChangeAttribution,
  ChangeSet,
  RevisionMutationResult,
  RevisionPage,
  RevisionRecord,
} from "../revisions/types.js";
import type {
  AgentDocxDocumentConfig,
  ResolveChangesInput,
} from "./contracts.js";
import {
  documentById,
  snapshotWithDependencies,
  snapshotWithSource,
  storedMediaType,
} from "./index.js";
import type { ProjectContext } from "./context.js";

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

export const applyRevisionDelta = (
  state: AttributionState,
  changes: readonly Change[],
  previousDocument?: LegalDocument,
  currentDocument?: LegalDocument,
): void => {
  const textChanges = changes.filter(
    (
      change,
    ): change is Extract<
      Change,
      { kind: "insert-text" | "delete-text" | "replace-text" }
    > =>
      change.kind === "insert-text" ||
      change.kind === "delete-text" ||
      change.kind === "replace-text",
  );
  for (const change of changes) {
    if ("blockId" in change) {
      if (
        change.kind === "insert-text" ||
        change.kind === "delete-text" ||
        change.kind === "replace-text"
      )
        continue;
      if (change.kind === "delete-block") {
        state.blocks.delete(change.blockId);
        state.operations.set(change.blockId, change.attribution);
      } else if ("newAttributionSpans" in change) {
        state.blocks.set(change.blockId, change.newAttributionSpans);
        state.operations.set(change.blockId, change.attribution);
      }
      continue;
    }
    if (change.kind === "add-config" || change.kind === "replace-config") {
      removeConfigAttribution(state.config, change.path);
      setConfigAttribution(
        state.config,
        change.newValue,
        change.path,
        change.attribution,
      );
      state.configOperations.set(change.path, change.attribution);
    } else if (change.kind === "remove-config") {
      removeConfigAttribution(state.config, change.path);
      state.configOperations.set(change.path, change.attribution);
    } else if (
      change.kind === "add-dependency" ||
      change.kind === "replace-dependency"
    ) {
      state.dependencies.set(change.key, change.attribution);
      state.dependencyOperations.set(change.key, change.attribution);
    } else if (change.kind === "remove-dependency") {
      state.dependencies.delete(change.key);
      state.dependencyOperations.set(change.key, change.attribution);
    }
  }
  if (!previousDocument || !currentDocument) return;
  const previousBlocks = new Map(
    provenanceBlocks(previousDocument).map((block) => [block.id, block]),
  );
  const currentBlocks = new Map(
    provenanceBlocks(currentDocument).map((block) => [block.id, block]),
  );
  for (const blockId of new Set(textChanges.map((change) => change.blockId))) {
    const previousBlock = previousBlocks.get(blockId);
    const currentBlock = currentBlocks.get(blockId);
    if (!previousBlock || !currentBlock)
      throw new AgentDocxError(
        "PROJECT_INVALID",
        `Text delta references a missing block: ${blockId}`,
      );
    const operation = textChanges.at(-1)!;
    state.blocks.set(
      blockId,
      reattributeVisibleText(
        visibleTextForBlock(previousBlock),
        visibleTextForBlock(currentBlock),
        state.blocks.get(blockId),
        operation.attribution,
      ),
    );
    state.operations.set(blockId, operation.attribution);
  }
  const sourceChanges = changes.filter(
    (change): change is Extract<Change, { blockId: string }> =>
      "blockId" in change,
  );
  const sourceOperation = sourceChanges.at(-1);
  if (!sourceOperation) return;
  const parentMap = (
    document: LegalDocument,
  ): {
    blocks: Map<string, AddressableBlock>;
    parents: Map<string, string | null>;
  } => {
    const blocks = new Map<string, AddressableBlock>();
    const parents = new Map<string, string | null>();
    const visit = (
      entries: readonly AddressableBlock[],
      parent: string | null,
    ): void => {
      for (const block of entries) {
        blocks.set(block.id, block);
        parents.set(block.id, parent);
        if (block.kind === "exhibit" || block.kind === "length-exclusion")
          visit(block.blocks, block.id);
        else if (block.kind === "list")
          for (const item of block.items) visit(item.children, block.id);
      }
    };
    visit(document.blocks, null);
    for (const footnote of document.footnotes) {
      blocks.set(footnote.id, footnote);
      parents.set(footnote.id, null);
    }
    return { blocks, parents };
  };
  const previousTree = parentMap(previousDocument);
  const currentTree = parentMap(currentDocument);
  const affectedContainers = new Set<string>();
  const addAncestors = (
    blockId: string,
    tree: ReturnType<typeof parentMap>,
  ): void => {
    let parentId = tree.parents.get(blockId) ?? null;
    while (parentId !== null) {
      const parent = tree.blocks.get(parentId);
      if (
        parent &&
        (parent.kind === "list" ||
          parent.kind === "exhibit" ||
          parent.kind === "length-exclusion")
      )
        affectedContainers.add(parentId);
      parentId = tree.parents.get(parentId) ?? null;
    }
  };
  for (const change of sourceChanges) {
    addAncestors(change.blockId, previousTree);
    addAncestors(change.blockId, currentTree);
  }
  for (const blockId of affectedContainers) {
    const previousBlock = previousTree.blocks.get(blockId);
    const currentBlock = currentTree.blocks.get(blockId);
    if (!previousBlock || !currentBlock) continue;
    state.blocks.set(
      blockId,
      reattributeVisibleText(
        visibleTextForBlock(previousBlock),
        visibleTextForBlock(currentBlock),
        state.blocks.get(blockId),
        sourceOperation.attribution,
      ),
    );
    state.operations.set(blockId, sourceOperation.attribution);
  }
};

export const applyRejectedConfigChanges = (
  head: AgentDocxDocumentConfig,
  changes: readonly Change[],
  decisions: Readonly<Record<`c_${string}`, "accept" | "reject">>,
): AgentDocxDocumentConfig => {
  const result = JSON.parse(canonicalJson(head)) as MutableJsonObject;
  for (const change of changes) {
    if (decisions[change.id] !== "reject") continue;
    if (change.kind === "add-config") {
      const { parent, key } = configParent(result, change.path);
      const current = parent[key];
      if (
        !Object.hasOwn(parent, key) ||
        canonicalJson(current) !== canonicalJson(change.newValue)
      )
        throw new AgentDocxError(
          "CHANGESET_INVALID",
          `Configuration add does not match head: ${change.path}`,
        );
      delete parent[key];
    } else if (change.kind === "remove-config") {
      const { parent, key } = configParent(result, change.path);
      if (Object.hasOwn(parent, key))
        throw new AgentDocxError(
          "CHANGESET_INVALID",
          `Configuration remove does not match head: ${change.path}`,
        );
      parent[key] = change.oldValue;
    } else if (change.kind === "replace-config") {
      const { parent, key } = configParent(result, change.path);
      const current = parent[key];
      if (
        !Object.hasOwn(parent, key) ||
        canonicalJson(current) !== canonicalJson(change.newValue)
      )
        throw new AgentDocxError(
          "CHANGESET_INVALID",
          `Configuration replacement does not match head: ${change.path}`,
        );
      parent[key] = change.oldValue;
    }
  }
  return result as AgentDocxDocumentConfig;
};

export const applyRejectedDependencyChanges = (
  head: Readonly<Record<string, RevisionId>>,
  changes: readonly Change[],
  decisions: Readonly<Record<`c_${string}`, "accept" | "reject">>,
): Record<string, RevisionId> => {
  const result = { ...head };
  for (const change of changes) {
    if (decisions[change.id] !== "reject") continue;
    if (change.kind === "add-dependency") {
      if (result[change.key] !== change.newObject)
        throw new AgentDocxError(
          "CHANGESET_INVALID",
          `Dependency add does not match head: ${change.key}`,
        );
      delete result[change.key];
    } else if (change.kind === "remove-dependency") {
      if (result[change.key] !== undefined)
        throw new AgentDocxError(
          "CHANGESET_INVALID",
          `Dependency remove does not match head: ${change.key}`,
        );
      result[change.key] = change.oldObject;
    } else if (change.kind === "replace-dependency") {
      if (result[change.key] !== change.newObject)
        throw new AgentDocxError(
          "CHANGESET_INVALID",
          `Dependency replacement does not match head: ${change.key}`,
        );
      result[change.key] = change.oldObject;
    }
  }
  return Object.fromEntries(
    Object.entries(result).sort(([left], [right]) =>
      left < right ? -1 : left > right ? 1 : 0,
    ),
  ) as Record<string, RevisionId>;
};

export const dependencyPath = (
  projectDirectory: string,
  config: AgentDocxDocumentConfig,
  key: string,
): string | null => {
  if (key === "template")
    return config.template ? resolve(projectDirectory, config.template) : null;
  if (key.startsWith("asset/")) {
    if (!config.assetsDir) return null;
    return resolve(
      projectDirectory,
      config.assetsDir,
      key.slice("asset/".length),
    );
  }
  if (key.startsWith("font/")) {
    const role = key.slice("font/".length) as
      | "regular"
      | "bold"
      | "italic"
      | "boldItalic";
    const configured = {
      regular: config.fontSet?.regularPath,
      bold: config.fontSet?.boldPath,
      italic: config.fontSet?.italicPath,
      boldItalic: config.fontSet?.boldItalicPath,
    }[role];
    return configured ? resolve(projectDirectory, configured) : null;
  }
  const customRulePack = /^rule-pack:(\d+)$/.exec(key);
  if (customRulePack) {
    const configured = config.rulePacks?.[Number(customRulePack[1])];
    return configured ? resolve(projectDirectory, configured) : null;
  }
  return null;
};

export const dependencyPathsChanged = (
  projectDirectory: string,
  currentConfig: AgentDocxDocumentConfig,
  targetConfig: AgentDocxDocumentConfig,
  currentDependencies: Readonly<Record<string, RevisionId>>,
  targetDependencies: Readonly<Record<string, RevisionId>>,
): boolean =>
  [
    ...new Set([
      ...Object.keys(currentDependencies),
      ...Object.keys(targetDependencies),
    ]),
  ].some(
    (key) =>
      dependencyPath(projectDirectory, currentConfig, key) !==
      dependencyPath(projectDirectory, targetConfig, key),
  );

export const materializeSelectedDependencies = async (
  opened: OpenedStore,
  currentConfig: AgentDocxDocumentConfig,
  targetConfig: AgentDocxDocumentConfig,
  currentDependencies: Readonly<Record<string, RevisionId>>,
  targetDependencies: Readonly<Record<string, RevisionId>>,
): Promise<void> => {
  const keys = [
    ...new Set([
      ...Object.keys(currentDependencies),
      ...Object.keys(targetDependencies),
    ]),
  ].sort((left, right) => (left < right ? -1 : left > right ? 1 : 0));
  const targetPaths = new Map<string, RevisionId>();
  const expectedIds = new Map<string, Set<RevisionId>>();
  const expectedId = (
    path: string | null,
    id: RevisionId | undefined,
  ): void => {
    if (path === null || id === undefined) return;
    const ids = expectedIds.get(path) ?? new Set<RevisionId>();
    ids.add(id);
    expectedIds.set(path, ids);
  };
  for (const key of keys) {
    const currentPath = dependencyPath(
      opened.projectDirectory,
      currentConfig,
      key,
    );
    const targetPath = dependencyPath(
      opened.projectDirectory,
      targetConfig,
      key,
    );
    expectedId(currentPath, currentDependencies[key]);
    expectedId(targetPath, targetDependencies[key]);
    const targetObject = targetDependencies[key];
    if (targetObject !== undefined && targetPath !== null) {
      const previous = targetPaths.get(targetPath);
      if (previous !== undefined && previous !== targetObject)
        throw new AgentDocxError(
          "PROJECT_INVALID",
          `Multiple dependencies target the same path: ${targetPath}`,
        );
      targetPaths.set(targetPath, targetObject);
    } else if (
      targetObject !== undefined &&
      key !== "profile" &&
      key !== "rule-pack" &&
      !key.startsWith("rule-pack:") &&
      !key.startsWith("rule-source/")
    ) {
      throw new AgentDocxError(
        "PROJECT_INVALID",
        `Dependency has no configured path: ${key}`,
      );
    }
  }
  const paths = [
    ...new Set([...expectedIds.keys(), ...targetPaths.keys()]),
  ].sort((left, right) => (left < right ? -1 : left > right ? 1 : 0));
  const states = new Map<
    string,
    { bytes: Uint8Array | null; id: RevisionId | null }
  >();
  const knownIds = new Set<RevisionId>([
    ...Object.values(currentDependencies),
    ...Object.values(targetDependencies),
  ]);
  for (const path of paths) {
    try {
      const entry = await lstat(path);
      if (!entry.isFile() || entry.isSymbolicLink())
        throw new AgentDocxError(
          "WORKING_COPY_CONFLICT",
          `Owned dependency is not a regular file: ${path}`,
        );
      const bytes = await readProjectFile(
        path,
        "Owned dependency",
        opened.projectDirectory,
      );
      const id = objectId(bytes);
      if (!expectedIds.get(path)?.has(id))
        throw new AgentDocxError(
          "WORKING_COPY_CONFLICT",
          `Owned dependency changed: ${path}`,
        );
      states.set(path, { bytes, id });
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code !== "ENOENT") throw error;
      states.set(path, { bytes: null, id: null });
    }
  }
  const restore = async (): Promise<void> => {
    const restorePaths = [...paths].sort(
      (left, right) =>
        right.length - left.length ||
        (left < right ? -1 : left > right ? 1 : 0),
    );
    for (const path of restorePaths) {
      const original = states.get(path)!;
      let currentBytes: Uint8Array | null = null;
      try {
        const entry = await lstat(path);
        if (!entry.isFile() || entry.isSymbolicLink())
          throw new AgentDocxError(
            "PROJECT_INVALID",
            `Cannot roll back non-regular dependency: ${path}`,
          );
        currentBytes = await readProjectFile(
          path,
          "Owned dependency",
          opened.projectDirectory,
        );
      } catch (error) {
        const code =
          error instanceof AgentDocxError
            ? error.code
            : (error as NodeJS.ErrnoException).code;
        if (code !== "INPUT_NOT_FOUND" && code !== "ENOENT") throw error;
      }
      if (original.bytes === null) {
        if (currentBytes === null) continue;
        const currentId = objectId(currentBytes);
        if (!knownIds.has(currentId))
          throw new AgentDocxError(
            "WORKING_COPY_CONFLICT",
            `Dependency changed during rollback: ${path}`,
          );
        await removeOwnedFile(path, currentId);
        continue;
      }
      if (currentBytes === null) {
        await mkdir(dirname(path), { recursive: true, mode: 0o700 });
        await replaceOwnedFile(path, null, original.bytes);
        continue;
      }
      const currentId = objectId(currentBytes);
      if (currentId === original.id) continue;
      if (!knownIds.has(currentId))
        throw new AgentDocxError(
          "WORKING_COPY_CONFLICT",
          `Dependency changed during rollback: ${path}`,
        );
      await replaceOwnedFile(path, currentId, original.bytes);
    }
  };
  const removeCurrent = async (path: string): Promise<void> => {
    const state = states.get(path)!;
    if (state.id !== null) await removeOwnedFile(path, state.id);
  };
  const overlappingOldPaths = paths
    .filter(
      (path) =>
        !targetPaths.has(path) &&
        [...targetPaths.keys()].some((targetPath) =>
          pathsOverlap(path, targetPath),
        ),
    )
    .sort(
      (left, right) =>
        right.length - left.length ||
        (left < right ? -1 : left > right ? 1 : 0),
    );
  try {
    for (const path of overlappingOldPaths) await removeCurrent(path);
    for (const [targetPath, targetObject] of [...targetPaths.entries()].sort(
      ([left], [right]) => (left < right ? -1 : left > right ? 1 : 0),
    )) {
      const state = states.get(targetPath)!;
      if (state.id === targetObject) continue;
      const bytes = await readObject(opened.storePath, targetObject);
      await mkdir(dirname(targetPath), { recursive: true, mode: 0o700 });
      await replaceOwnedFile(targetPath, state.id, bytes);
    }
    for (const path of paths)
      if (!targetPaths.has(path) && !overlappingOldPaths.includes(path))
        await removeCurrent(path);
  } catch (error) {
    await restore();
    throw error;
  }
};

export const rejectedSourceReplacements = (
  source: string,
  baseSource: string,
  changes: readonly Change[],
  decisions: Readonly<Record<`c_${string}`, "accept" | "reject">>,
): RawReplacement[] => {
  const sourceChanges = changes.filter(
    (change) =>
      change.kind !== "add-config" &&
      change.kind !== "remove-config" &&
      change.kind !== "replace-config" &&
      change.kind !== "add-dependency" &&
      change.kind !== "remove-dependency" &&
      change.kind !== "replace-dependency",
  );
  if (
    sourceChanges.length > 0 &&
    sourceChanges.every((change) => decisions[change.id] === "reject")
  )
    return [
      {
        start: 0,
        end: source.length,
        expectedText: source,
        replacement: baseSource,
      },
    ];
  const replacements: RawReplacement[] = [];
  for (const change of changes) {
    if (decisions[change.id] !== "reject") continue;
    if (change.kind === "insert-block" || change.kind === "insert-text") {
      const range =
        change.kind === "insert-block"
          ? sourceRangeWithMarker(
              source,
              change.newSource.start,
              change.newSource.end,
              change.blockId,
            )
          : change.newSource;
      replacements.push({
        start: range.start,
        end: range.end,
        expectedText: source.slice(range.start, range.end),
        replacement: "",
      });
      continue;
    }
    if (change.kind === "delete-block") {
      const start = sourceInsertionOffset(
        baseSource,
        source,
        change.oldSource.start,
        change.oldSource.end,
      );
      replacements.push({
        start,
        end: start,
        expectedText: "",
        replacement: sourceInsertionText(
          baseSource,
          change.oldSource.start,
          change.oldSource.end,
        ),
      });
      continue;
    }
    if (change.kind === "move-block") {
      const range = change.newSource;
      const insertion = sourceInsertionOffset(
        baseSource,
        source,
        change.oldSource.start,
        change.oldSource.end,
      );
      replacements.push({
        start: range.start,
        end: range.end,
        expectedText: range.text,
        replacement: "",
      });
      replacements.push({
        start: insertion,
        end: insertion,
        expectedText: "",
        replacement: sourceInsertionText(
          baseSource,
          change.oldSource.start,
          change.oldSource.end,
        ),
      });
      continue;
    }
    if (change.kind === "replace-container-shell") {
      if (
        change.oldShell.sourceRanges.length !==
        change.newShell.sourceRanges.length
      )
        throw new AgentDocxError(
          "CHANGESET_INVALID",
          "Container shell source ranges do not align",
        );
      for (const [index, range] of change.newShell.sourceRanges.entries()) {
        const oldRange = change.oldShell.sourceRanges[index]!;
        replacements.push({
          start: range.start,
          end: range.end,
          expectedText: range.text,
          replacement: oldRange.text,
        });
      }
      continue;
    }
    if (change.kind === "delete-text") {
      replacements.push({
        start: change.newOffset,
        end: change.newOffset,
        expectedText: "",
        replacement: change.oldSource.text,
      });
      continue;
    }
    if (change.kind === "replace-text") {
      replacements.push({
        start: change.newSource.start,
        end: change.newSource.end,
        expectedText: change.newSource.text,
        replacement: change.oldSource.text,
      });
      continue;
    }
    if (change.kind === "replace-block") {
      replacements.push({
        start: change.newBlock.position.start.offset,
        end: change.newBlock.position.end.offset,
        expectedText: change.newBlock.sourceText,
        replacement: change.oldBlock.sourceText,
      });
      continue;
    }
    if (
      change.kind === "add-config" ||
      change.kind === "remove-config" ||
      change.kind === "replace-config" ||
      change.kind === "add-dependency" ||
      change.kind === "remove-dependency" ||
      change.kind === "replace-dependency"
    )
      continue;
    throw new AgentDocxError(
      "CHANGESET_INVALID",
      "Cannot safely reject a change without an exact head source range",
    );
  }
  for (const replacement of replacements)
    if (
      source.slice(replacement.start, replacement.end) !==
      replacement.expectedText
    )
      throw new AgentDocxError(
        "REVISION_CONFLICT",
        "Change-set head source no longer matches its recorded range",
      );
  const ordered = [...replacements].sort(
    (left, right) => right.start - left.start,
  );
  for (const [index, replacement] of ordered.entries()) {
    const next = ordered[index + 1];
    if (next && replacement.start < next.end && next.start < replacement.end)
      throw new AgentDocxError(
        "CHANGESET_INVALID",
        "Rejected source changes overlap",
      );
  }
  return ordered;
};

export const sourceRangeWithMarker = (
  source: string,
  start: number,
  end: number,
  blockId: string,
): { start: number; end: number } => {
  const marker = `<!-- agent-docx:block id="${blockId}" -->`;
  const markerStart = source.lastIndexOf(marker, start);
  if (
    markerStart < 0 ||
    (markerStart !== start &&
      !/^\r?\n$/.test(source.slice(markerStart + marker.length, start)))
  )
    return { start, end };
  const nextMarker = source.indexOf('<!-- agent-docx:block id="', end);
  if (nextMarker >= 0 && /^[\r\n]*$/.test(source.slice(end, nextMarker)))
    return { start: markerStart, end: nextMarker };
  const newlineLength = source.startsWith("\r\n", end)
    ? 2
    : source[end] === "\n"
      ? 1
      : 0;
  return {
    start: markerStart,
    end: end + newlineLength,
  };
};

export const sourceMarkerLines = (
  source: string,
): readonly SourceMarkerLine[] => {
  const markers: SourceMarkerLine[] = [];
  const pattern = /<!--[ \t]*agent-docx:block[ \t]+id="([^"]+)"[ \t]*-->/g;
  for (const match of source.matchAll(pattern)) {
    const markerStart = match.index ?? 0;
    const lineStart = source.lastIndexOf("\n", markerStart - 1) + 1;
    markers.push({
      id: match[1]!,
      start: lineStart,
      end: markerStart + match[0].length,
    });
  }
  return markers;
};

export const sourceInsertionOffset = (
  baseSource: string,
  headSource: string,
  baseStart: number,
  baseEnd: number,
): number => {
  const baseMarkers = sourceMarkerLines(baseSource);
  const headMarkers = new Map(
    sourceMarkerLines(headSource).map((marker) => [marker.id, marker]),
  );
  const previous = [...baseMarkers]
    .reverse()
    .find((marker) => marker.end <= baseStart);
  const next = baseMarkers.find((marker) => marker.start >= baseEnd);
  const headNext = next ? headMarkers.get(next.id) : undefined;
  if (headNext) return headNext.start;
  const headPrevious = previous ? headMarkers.get(previous.id) : undefined;
  if (headPrevious) {
    const following = sourceMarkerLines(headSource).find(
      (marker) => marker.start > headPrevious.start,
    );
    return following?.start ?? headSource.length;
  }
  return 0;
};

export const sourceInsertionText = (
  baseSource: string,
  baseStart: number,
  baseEnd: number,
): string => {
  const oldText = baseSource.slice(baseStart, baseEnd);
  const markers = sourceMarkerLines(baseSource);
  const next = markers.find((marker) => marker.start >= baseEnd);
  if (next) return oldText + baseSource.slice(baseEnd, next.start);
  return oldText + baseSource.slice(baseEnd);
};

export const jsonPointerParts = (path: string): readonly string[] => {
  if (path === "") return [];
  if (!path.startsWith("/"))
    throw new AgentDocxError(
      "CHANGESET_INVALID",
      `Invalid configuration path: ${path}`,
    );
  return path
    .slice(1)
    .split("/")
    .map((part) => part.replaceAll("~1", "/").replaceAll("~0", "~"));
};

export const configParent = (
  root: MutableJsonObject,
  path: string,
): { parent: MutableJsonObject; key: string } => {
  const parts = jsonPointerParts(path);
  if (parts.length === 0)
    throw new AgentDocxError(
      "CHANGESET_INVALID",
      "Root configuration replacement is not supported",
    );
  let parent: MutableJsonObject = root;
  for (const part of parts.slice(0, -1)) {
    const child = parent[part];
    if (child === null || typeof child !== "object" || Array.isArray(child))
      throw new AgentDocxError(
        "CHANGESET_INVALID",
        `Configuration path is missing: ${path}`,
      );
    parent = child as MutableJsonObject;
  }
  return { parent, key: parts.at(-1)! };
};

export const isUtf16Boundary = (text: string, offset: number): boolean =>
  Number.isInteger(offset) &&
  offset >= 0 &&
  offset <= text.length &&
  (offset === 0 ||
    offset === text.length ||
    !(
      text.charCodeAt(offset - 1) >= 0xd800 &&
      text.charCodeAt(offset - 1) <= 0xdbff &&
      text.charCodeAt(offset) >= 0xdc00 &&
      text.charCodeAt(offset) <= 0xdfff
    ));

export const currentRevision = async (
  opened: OpenedStore,
  documentId: string,
  selector: RevisionId | "HEAD",
): Promise<RevisionRecord> => {
  const head = await readHead(opened.storePath, documentId);
  const requested = selector === "HEAD" ? head : selector;
  if (!requested)
    throw new AgentDocxError(
      "REVISION_NOT_FOUND",
      `Document has no revision: ${documentId}`,
    );
  const visited = new Set<RevisionId>();
  const pending: {
    id: RevisionId;
    ancestry: ReadonlySet<RevisionId>;
  }[] = head ? [{ id: head, ancestry: new Set() }] : [];
  while (pending.length > 0) {
    const entry = pending.pop()!;
    if (entry.ancestry.has(entry.id))
      throw new AgentDocxError(
        "PROJECT_INVALID",
        "Revision graph contains a cycle",
      );
    if (visited.has(entry.id)) continue;
    visited.add(entry.id);
    const record = await readRevisionJson<RevisionRecord>(
      opened.storePath,
      entry.id,
    );
    if (record.documentId !== documentId)
      throw new AgentDocxError(
        "PROJECT_INVALID",
        `Revision belongs to another document: ${entry.id}`,
      );
    if (record.id === requested) return record;
    const ancestry = new Set(entry.ancestry);
    ancestry.add(entry.id);
    pending.push(...record.parents.map((id) => ({ id, ancestry })));
  }
  throw new AgentDocxError(
    "REVISION_NOT_FOUND",
    `Revision not found: ${requested}`,
  );
};

export const isFirstParentAncestor = async (
  opened: OpenedStore,
  ancestor: RevisionId,
  descendant: RevisionRecord,
): Promise<boolean> => {
  const visited = new Set<RevisionId>();
  let current: RevisionRecord | null = descendant;
  while (current) {
    if (visited.has(current.id))
      throw new AgentDocxError(
        "PROJECT_INVALID",
        "Revision graph contains a cycle",
      );
    visited.add(current.id);
    if (current.documentId !== descendant.documentId)
      throw new AgentDocxError(
        "PROJECT_INVALID",
        "Revision graph crosses documents",
      );
    if (current.id === ancestor) return true;
    const parent: RevisionId | undefined = current.parents[0];
    current = parent
      ? await readRevisionJson<RevisionRecord>(opened.storePath, parent)
      : null;
  }
  return false;
};

export const materialFor = async (
  opened: OpenedStore,
  record: RevisionRecord,
): Promise<RevisionMaterial> => {
  const source = new TextDecoder("utf-8", { fatal: true }).decode(
    await readObject(opened.storePath, record.sourceObject),
  );
  const config = assertStoredConfig(
    strictJson(
      await readObject(opened.storePath, record.documentConfigObject),
      record.documentConfigObject,
    ),
    record.documentConfigObject,
  );
  const document = assertStoredDocument(
    strictJson(
      await readObject(opened.storePath, record.legalDocumentObject),
      record.legalDocumentObject,
    ),
    record.legalDocumentObject,
  );
  const annotations = assertStoredAnnotations(
    strictJson(
      await readObject(opened.storePath, record.annotationsObject),
      record.annotationsObject,
    ),
    record.annotationsObject,
  );
  return { revision: record, source, config, document, annotations };
};

export const snapshotForMaterial = async (
  opened: OpenedStore,
  material: RevisionMaterial,
): Promise<ProjectSnapshot> => {
  const dependencyBytes = new Map<
    string,
    { bytes: Uint8Array; mediaType: string }
  >();
  for (const [key, id] of Object.entries(material.revision.dependencyObjects))
    dependencyBytes.set(key, {
      bytes: await readObject(opened.storePath, id),
      mediaType: storedMediaType(key),
    });
  return {
    source: material.source,
    sourceObject: material.revision.sourceObject,
    documentConfigObject: material.revision.documentConfigObject,
    dependencyObjects: material.revision.dependencyObjects,
    dependencyBytes,
    workingTreeHash: material.revision.workingTreeHash,
  };
};

export const annotationsForHead = async (
  opened: OpenedStore,
  documentId: string,
): Promise<readonly ReviewAnnotation[]> => {
  const head = await readHead(opened.storePath, documentId);
  if (!head) return [];
  const record = await readRevisionJson<RevisionRecord>(opened.storePath, head);
  return assertStoredAnnotations(
    strictJson(
      await readObject(opened.storePath, record.annotationsObject),
      record.annotationsObject,
    ),
    record.annotationsObject,
  );
};

export const listRevisions = async (
  ctx: ProjectContext,
  documentId: string,
  input: { limit?: number; cursor?: RevisionId } = {},
): Promise<RevisionPage> => {
  const limit = input.limit ?? 100;
  if (!Number.isInteger(limit) || limit < 1 || limit > 1000)
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      "Revision limit must be 1 through 1000",
    );
  return withLockedStore(ctx.manifestPath, async (opened) => {
    const head = await readHead(opened.storePath, documentId);
    const reachable = new Map<RevisionId, RevisionRecord>();
    const pending: {
      id: RevisionId;
      ancestry: ReadonlySet<RevisionId>;
    }[] = head ? [{ id: head, ancestry: new Set() }] : [];
    let cursorFound = input.cursor === undefined;
    while (pending.length > 0) {
      const entry = pending.pop()!;
      if (entry.ancestry.has(entry.id))
        throw new AgentDocxError(
          "PROJECT_INVALID",
          "Revision graph contains a cycle",
        );
      if (reachable.has(entry.id)) continue;
      const record = await readRevisionJson<RevisionRecord>(
        opened.storePath,
        entry.id,
      );
      if (record.documentId !== documentId)
        throw new AgentDocxError(
          "PROJECT_INVALID",
          "Revision graph crosses documents",
        );
      reachable.set(entry.id, record);
      if (record.id === input.cursor) cursorFound = true;
      const ancestry = new Set(entry.ancestry);
      ancestry.add(entry.id);
      pending.push(...record.parents.map((id) => ({ id, ancestry })));
      const ordered = [...reachable.values()].sort(
        (left, right) =>
          (right.createdAt < left.createdAt
            ? -1
            : right.createdAt > left.createdAt
              ? 1
              : 0) || (left.id < right.id ? -1 : left.id > right.id ? 1 : 0),
      );
      if (input.cursor === undefined && ordered.length >= limit + 1) break;
      if (cursorFound && input.cursor !== undefined) {
        const cursorIndex = ordered.findIndex(
          (candidate) => candidate.id === input.cursor,
        );
        if (cursorIndex >= 0 && ordered.length >= cursorIndex + limit + 2)
          break;
      }
    }
    const ordered = [...reachable.values()].sort(
      (left, right) =>
        (right.createdAt < left.createdAt
          ? -1
          : right.createdAt > left.createdAt
            ? 1
            : 0) || (left.id < right.id ? -1 : left.id > right.id ? 1 : 0),
    );
    const start = input.cursor
      ? Math.max(
          0,
          ordered.findIndex((record) => record.id === input.cursor) + 1,
        )
      : 0;
    if (input.cursor && start === 0)
      throw new AgentDocxError(
        "REVISION_NOT_FOUND",
        `Revision cursor not found: ${input.cursor}`,
      );
    const items = ordered.slice(start, start + limit);
    return {
      schemaVersion: 1,
      items,
      nextCursor: ordered[start + limit]?.id ?? null,
    };
  });
};

export const getRevision = async (
  ctx: ProjectContext,
  documentId: string,
  revision: RevisionId | "HEAD",
): Promise<RevisionRecord> =>
  withLockedStore(ctx.manifestPath, async (opened) =>
    currentRevision(opened, documentId, revision),
  );

export const diff = async (
  ctx: ProjectContext,
  documentId: string,
  base: RevisionId | "HEAD",
  head: RevisionId | "HEAD",
): Promise<ChangeSet> =>
  withLockedStore(ctx.manifestPath, async (opened) => {
    const baseRecord = await currentRevision(opened, documentId, base);
    const headRecord = await currentRevision(opened, documentId, head);
    if (
      baseRecord.id !== headRecord.id &&
      !(await isFirstParentAncestor(opened, baseRecord.id, headRecord))
    )
      throw new AgentDocxError(
        "REVISION_CONFLICT",
        "Diff base must be a first-parent ancestor of head",
      );
    const baseMaterial = await materialFor(opened, baseRecord);
    const headMaterial = await materialFor(opened, headRecord);
    const changeSet = createChangeSet(
      documentId,
      baseRecord.id,
      headRecord.id,
      baseMaterial.document,
      headMaterial.document,
      baseMaterial.annotations,
      headMaterial.annotations,
      defaultAttribution(headRecord.author, headRecord.createdAt),
      {
        baseConfig: baseMaterial.config as unknown as JsonObject,
        headConfig: headMaterial.config as unknown as JsonObject,
        baseDependencies: baseRecord.dependencyObjects,
        headDependencies: headRecord.dependencyObjects,
        baseSource: baseMaterial.source,
        headSource: headMaterial.source,
      },
    );
    const baseProvenance = await provenanceForRevision(opened, baseRecord);
    const headProvenance = await provenanceForRevision(opened, headRecord);
    const provenance: ChangeSetProvenance = {
      baseBlocks: baseProvenance.blocks,
      headBlocks: headProvenance.blocks,
      baseOperations: baseProvenance.operations,
      headOperations: headProvenance.operations,
      baseConfig: baseProvenance.config,
      headConfig: headProvenance.config,
      baseDependencies: baseProvenance.dependencies,
      headDependencies: headProvenance.dependencies,
      baseConfigOperations: baseProvenance.configOperations,
      headConfigOperations: headProvenance.configOperations,
      baseDependencyOperations: baseProvenance.dependencyOperations,
      headDependencyOperations: headProvenance.dependencyOperations,
      baseDocument: baseMaterial.document,
      headDocument: headMaterial.document,
    };
    return reattributeChangeSet(changeSet, provenance);
  });

export const restore = async (
  ctx: ProjectContext,
  documentId: string,
  input: {
    baseRevision: RevisionId | "HEAD";
    targetRevision: RevisionId | "HEAD";
    author: Actor;
    message: string;
  },
): Promise<RevisionMutationResult> => {
  return withLockedStore(ctx.manifestPath, async (opened) => {
    const base = await currentRevision(opened, documentId, input.baseRevision);
    const target = await currentRevision(
      opened,
      documentId,
      input.targetRevision,
    );
    if ((await readHead(opened.storePath, documentId)) !== base.id)
      throw new AgentDocxError(
        "REVISION_CONFLICT",
        "Restore base is not the current head",
      );
    const currentConfig = documentById(opened.manifest, documentId);
    const currentSnapshot = await snapshotProjectDocument(
      opened,
      currentConfig,
    );
    if (currentSnapshot.workingTreeHash !== base.workingTreeHash)
      throw new AgentDocxError(
        "WORKING_COPY_CONFLICT",
        "Working copy differs from restore base",
      );
    const material = await materialFor(opened, target);
    if (material.config.source !== currentConfig.source)
      throw new AgentDocxError(
        "PROJECT_INVALID",
        "Restore cannot change the document source path",
      );
    const targetSnapshot = await snapshotWithDependencies(
      opened,
      currentSnapshot,
      material.config,
      target.dependencyObjects,
    );
    const snapshot = snapshotWithSource(targetSnapshot, material.source);
    return commitLocked(
      ctx,
      opened,
      material.config,
      snapshot,
      material.document,
      material.annotations,
      base.id,
      input.author,
      input.message,
      undefined,
      true,
      true,
      { expectedWorkingTreeHash: currentSnapshot.workingTreeHash },
    );
  });
};

export const resolveChanges = async (
  ctx: ProjectContext,
  documentId: string,
  input: ResolveChangesInput,
): Promise<RevisionMutationResult> => {
  return withLockedStore(ctx.manifestPath, async (opened) => {
    const base = await currentRevision(
      opened,
      documentId,
      input.changeSet.baseRevision,
    );
    const head = await currentRevision(
      opened,
      documentId,
      input.changeSet.headRevision,
    );
    if (
      base.id === head.id ||
      !(await isFirstParentAncestor(opened, base.id, head))
    )
      throw new AgentDocxError(
        "CHANGESET_INVALID",
        "Change-set base must be a distinct first-parent ancestor",
      );
    const baseMaterial = await materialFor(opened, base);
    const headMaterial = await materialFor(opened, head);
    const rawExpected = createChangeSet(
      documentId,
      base.id,
      head.id,
      baseMaterial.document,
      headMaterial.document,
      baseMaterial.annotations,
      headMaterial.annotations,
      defaultAttribution(head.author, head.createdAt),
      {
        baseConfig: baseMaterial.config as unknown as JsonObject,
        headConfig: headMaterial.config as unknown as JsonObject,
        baseDependencies: base.dependencyObjects,
        headDependencies: head.dependencyObjects,
        baseSource: baseMaterial.source,
        headSource: headMaterial.source,
      },
    );
    const baseProvenance = await provenanceForRevision(opened, base);
    const headProvenance = await provenanceForRevision(opened, head);
    const expected = reattributeChangeSet(rawExpected, {
      baseBlocks: baseProvenance.blocks,
      headBlocks: headProvenance.blocks,
      baseOperations: baseProvenance.operations,
      headOperations: headProvenance.operations,
      baseConfig: baseProvenance.config,
      headConfig: headProvenance.config,
      baseConfigOperations: baseProvenance.configOperations,
      headConfigOperations: headProvenance.configOperations,
      baseDependencyOperations: baseProvenance.dependencyOperations,
      headDependencyOperations: headProvenance.dependencyOperations,
      baseDocument: baseMaterial.document,
      headDocument: headMaterial.document,
      baseDependencies: baseProvenance.dependencies,
      headDependencies: headProvenance.dependencies,
    });
    if (canonicalJson(expected) !== canonicalJson(input.changeSet))
      throw new AgentDocxError(
        "CHANGESET_INVALID",
        "Change set does not match the selected immutable revisions",
      );
    if (
      base.id === head.id ||
      (await readHead(opened.storePath, documentId)) !== head.id
    )
      throw new AgentDocxError(
        "REVISION_CONFLICT",
        "Change-set head must be the distinct current document head",
      );
    const changeIds = [
      ...expected.changes.map((change) => change.id),
      ...expected.annotations.map((change) => change.id),
    ].sort((left, right) => (left < right ? -1 : left > right ? 1 : 0));
    const decisionIds = Object.keys(input.decisions).sort((left, right) =>
      left < right ? -1 : left > right ? 1 : 0,
    );
    if (
      Object.values(input.decisions).some(
        (decision) => decision !== "accept" && decision !== "reject",
      )
    )
      throw new AgentDocxError(
        "CHANGESET_INVALID",
        "Change-set decisions must be accept or reject",
      );
    if (
      changeIds.length !== decisionIds.length ||
      changeIds.some((id, index) => id !== decisionIds[index])
    )
      throw new AgentDocxError(
        "CHANGESET_INVALID",
        "Change-set decisions must select every change exactly once",
      );
    const currentConfig = documentById(opened.manifest, documentId);
    const snapshot = await snapshotProjectDocument(opened, currentConfig);
    if (
      snapshot.workingTreeHash !== head.workingTreeHash ||
      snapshot.sourceObject !== head.sourceObject ||
      snapshot.documentConfigObject !== head.documentConfigObject ||
      canonicalJson(snapshot.dependencyObjects) !==
        canonicalJson(head.dependencyObjects)
    )
      throw new AgentDocxError(
        "WORKING_COPY_CONFLICT",
        "Working copy differs from the change-set head",
      );
    const targetConfig = applyRejectedConfigChanges(
      headMaterial.config,
      expected.changes,
      input.decisions,
    );
    const targetDependencies = applyRejectedDependencyChanges(
      head.dependencyObjects,
      expected.changes,
      input.decisions,
    );
    const replacements = rejectedSourceReplacements(
      snapshot.source,
      baseMaterial.source,
      expected.changes,
      input.decisions,
    );
    let source = snapshot.source;
    for (const replacement of replacements)
      source = `${source.slice(0, replacement.start)}${replacement.replacement}${source.slice(replacement.end)}`;
    let annotations = [...headMaterial.annotations];
    for (const change of expected.annotations) {
      if (input.decisions[change.id] !== "reject") continue;
      if (change.kind === "add")
        annotations = annotations.filter(
          (annotation) => annotation.id !== change.newValue.id,
        );
      else if (change.kind === "remove") annotations.push(change.oldValue);
      else
        annotations = annotations.map((annotation) =>
          annotation.id === change.newValue.id ? change.oldValue : annotation,
        );
    }
    const targetSnapshot = await snapshotWithDependencies(
      opened,
      snapshot,
      targetConfig,
      targetDependencies,
    );
    const preparedSnapshot = snapshotWithSource(targetSnapshot, source);
    const document = documentFor(
      source,
      targetConfig,
      preparedSnapshot,
      opened.manifest.projectId,
      annotations,
      true,
    );
    return commitLocked(
      ctx,
      opened,
      targetConfig,
      preparedSnapshot,
      document,
      annotations,
      head.id,
      input.author,
      input.message,
      { schemaVersion: 1, changeSet: expected, decisions: input.decisions },
      false,
      true,
      {
        expectedWorkingTreeHash: snapshot.workingTreeHash,
        parentIds: [head.id, base.id],
        firstParent: head,
      },
    );
  });
};
