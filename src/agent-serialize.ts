import { isAbsolute } from "node:path";
import type { AgentAction } from "./agent-protocol.js";
import {
  isStatelessAgentRequest,
  revisionIdPattern,
} from "./agent-protocol.js";
import { publicPath } from "./path-util.js";
import type { RevisionId } from "./legal/model.js";
import type { JsonValue } from "./types.js";

export const responseMeta = (
  action: AgentAction,
  params: Record<string, unknown>,
  value: unknown,
): { documentId: string | null; revision: RevisionId | null } => {
  const patch =
    typeof params.patch === "object" &&
    params.patch !== null &&
    !Array.isArray(params.patch)
      ? (params.patch as Record<string, unknown>)
      : null;
  const documentId =
    typeof params.documentId === "string"
      ? params.documentId
      : patch && typeof patch.documentId === "string"
        ? patch.documentId
        : null;
  const record =
    typeof value === "object" && value !== null && !Array.isArray(value)
      ? (value as Record<string, unknown>)
      : null;
  const revisionId = (candidate: unknown): RevisionId | null =>
    typeof candidate === "string" && revisionIdPattern.test(candidate)
      ? (candidate as RevisionId)
      : null;

  if (action === "docx.inspect" || isStatelessAgentRequest(action, params))
    return { documentId: null, revision: null };

  if (action === "project.init" || action === "project.add") {
    const documents = record?.documents;
    const document = Array.isArray(documents)
      ? documents.find(
          (entry) =>
            typeof entry === "object" &&
            entry !== null &&
            !Array.isArray(entry) &&
            (entry as Record<string, unknown>).documentId === documentId,
        )
      : undefined;
    const head =
      typeof document === "object" &&
      document !== null &&
      !Array.isArray(document)
        ? revisionId((document as Record<string, unknown>).head)
        : null;
    return { documentId, revision: head };
  }

  if (action === "project.get") return { documentId: null, revision: null };

  if (action === "document.get")
    return {
      documentId,
      revision: record ? revisionId(record.revision) : null,
    };

  if (action === "revision.list") return { documentId, revision: null };

  if (action === "revision.get")
    return { documentId, revision: record ? revisionId(record.id) : null };

  if (action === "revision.diff") {
    const changeSet =
      record?.changeSet &&
      typeof record.changeSet === "object" &&
      !Array.isArray(record.changeSet)
        ? (record.changeSet as Record<string, unknown>)
        : null;
    return {
      documentId,
      revision: changeSet ? revisionId(changeSet.headRevision) : null,
    };
  }

  if (action === "draft.evaluate") {
    const patch =
      typeof params.patch === "object" &&
      params.patch !== null &&
      !Array.isArray(params.patch)
        ? (params.patch as Record<string, unknown>)
        : null;
    return {
      documentId:
        patch && typeof patch.documentId === "string" ? patch.documentId : null,
      revision: patch ? revisionId(patch.baseRevision) : null,
    };
  }

  if (action === "docx.import")
    return {
      documentId,
      revision: record ? revisionId(record.headRevision) : null,
    };

  if (action === "docx.importRedline")
    return {
      documentId,
      revision: record ? revisionId(record.headRevision) : null,
    };

  if (record) {
    const head = revisionId(record.head);
    if (head) return { documentId, revision: head };

    const revision =
      typeof record.revision === "object" &&
      record.revision !== null &&
      !Array.isArray(record.revision)
        ? (record.revision as Record<string, unknown>)
        : null;
    const mutationRevision = revision ? revisionId(revision.id) : null;
    if (mutationRevision) return { documentId, revision: mutationRevision };

    const selectedRevision = revisionId(record.revision);
    if (selectedRevision) return { documentId, revision: selectedRevision };

    const measurement =
      typeof record.measurement === "object" &&
      record.measurement !== null &&
      !Array.isArray(record.measurement)
        ? (record.measurement as Record<string, unknown>)
        : null;
    const measuredRevision = measurement
      ? revisionId(measurement.revision)
      : null;
    if (measuredRevision) return { documentId, revision: measuredRevision };

    const artifact =
      typeof record.artifact === "object" &&
      record.artifact !== null &&
      !Array.isArray(record.artifact)
        ? (record.artifact as Record<string, unknown>)
        : null;
    const artifactRevision = artifact ? revisionId(artifact.revision) : null;
    if (artifactRevision) return { documentId, revision: artifactRevision };
  }
  return { documentId, revision: null };
};

/**
 * Projects API values onto JSON-safe protocol values. Generated DOCX and
 * in-memory attachment payloads are intentionally omitted rather than encoded.
 */
export const serializeAgentValue = (
  value: unknown,
  cwd?: string,
): JsonValue => {
  const serialize = (
    candidate: unknown,
    key?: string,
  ): JsonValue | undefined => {
    if (candidate === undefined || candidate instanceof Uint8Array)
      return undefined;
    if (candidate === null) return null;
    if (typeof candidate === "string") {
      if (
        cwd &&
        key !== undefined &&
        ["manifestPath", "path", "storePath", "output"].includes(key) &&
        isAbsolute(candidate)
      )
        return publicPath(cwd, candidate);
      return candidate;
    }
    if (typeof candidate === "boolean") return candidate;
    if (typeof candidate === "number")
      return Number.isFinite(candidate) ? candidate : null;
    if (Array.isArray(candidate))
      return candidate.map((entry) => serialize(entry) ?? null);
    if (typeof candidate !== "object") return String(candidate);

    const output: Record<string, JsonValue> = {};
    for (const [childKey, child] of Object.entries(
      candidate as Record<string, unknown>,
    )) {
      if (childKey === "generatedDocx" || child instanceof Uint8Array) continue;
      if (
        childKey === "attachments" &&
        typeof child === "object" &&
        child !== null &&
        !Array.isArray(child) &&
        "files" in child
      )
        continue;
      const projected = serialize(child, childKey);
      if (projected !== undefined) output[childKey] = projected;
    }
    return output;
  };
  return serialize(value) ?? null;
};
