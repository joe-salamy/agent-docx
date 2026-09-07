import { AgentDocxError } from "../../types.js";
import {
  canonicalJson,
  readHead,
  readObject,
  readRevisionJson,
  snapshotProjectDocument,
  strictJson,
  withLockedStore,
  type OpenedStore,
  type ProjectSnapshot,
} from "../store.js";
import {
  assertStoredAnnotations,
  assertStoredConfig,
  assertStoredDocument,
  commitLocked,
  documentFor,
} from "../documents.js";
import {
  createChangeSet,
  defaultAttribution,
  reattributeChangeSet,
  type ChangeSetProvenance,
  type JsonObject,
} from "../../revisions/diff.js";
import {
  documentById,
  snapshotWithDependencies,
  snapshotWithSource,
  storedMediaType,
} from "../index.js";
import {
  provenanceForRevision,
  type MutableJsonObject,
  type RevisionMaterial,
  type SourceMarkerLine,
} from "./attribution.js";
import {
  applyRejectedConfigChanges,
  applyRejectedDependencyChanges,
  rejectedSourceReplacements,
} from "./deltas.js";
import type { ProjectContext } from "../context.js";
import type { Actor, ReviewAnnotation, RevisionId } from "../../legal/model.js";
import type {
  ChangeSet,
  RevisionMutationResult,
  RevisionPage,
  RevisionRecord,
} from "../../revisions/types.js";
import type { ResolveChangesInput } from "../contracts.js";

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
