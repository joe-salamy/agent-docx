import { AgentDocxError } from "../../types.js";
import {
  canonicalJson,
  objectId,
  readObject,
  readProjectFile,
  removeOwnedFile,
  replaceOwnedFile,
  type OpenedStore,
} from "../store.js";
import { reattributeVisibleText } from "../../revisions/diff.js";
import { visibleTextForBlock } from "../../legal/visible-text.js";
import { lstat, mkdir } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { pathsOverlap } from "../fs-util.js";
import {
  provenanceBlocks,
  removeConfigAttribution,
  setConfigAttribution,
  type AttributionState,
  type MutableJsonObject,
  type RawReplacement,
} from "./attribution.js";
import {
  configParent,
  sourceInsertionOffset,
  sourceInsertionText,
  sourceRangeWithMarker,
} from "./markers.js";
import type {
  AddressableBlock,
  LegalDocument,
  RevisionId,
} from "../../legal/model.js";
import type { Change } from "../../revisions/types.js";
import type { AgentDocxDocumentConfig } from "../contracts.js";

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
