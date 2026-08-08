import { AgentDocxError } from "../types.js";
import {
  readHead,
  readRevisionJson,
  snapshotProjectDocument,
  updateManifest,
  withLockedStore,
} from "./store.js";
import { getStateLocked, measureLocked, validateLocked } from "./documents.js";
import { documentById } from "./index.js";
import { isDocumentId } from "../legal/model.js";
import type {
  AgentDocxManifest,
  FilingSet,
  FilingSetSnapshot,
  FilingSetValidation,
  ProjectState,
} from "./contracts.js";
import type { RevisionRecord } from "../revisions/types.js";
import type { ValidationResult } from "../legal/rules.js";
import type { ProjectContext } from "./context.js";

export const addFilingSet = async (
  ctx: ProjectContext,
  input: {
    id: string;
    label?: string;
    documentIds: readonly string[];
    pageCap?: number;
  },
): Promise<ProjectState> => {
  return withLockedStore(ctx.manifestPath, async (opened) => {
    if (
      !input ||
      typeof input !== "object" ||
      Array.isArray(input) ||
      typeof input.id !== "string" ||
      !isDocumentId(input.id)
    )
      throw new AgentDocxError("INVALID_ARGUMENT", "Filing set id is invalid");
    if (
      input.label !== undefined &&
      (typeof input.label !== "string" || input.label.length === 0)
    )
      throw new AgentDocxError(
        "INVALID_ARGUMENT",
        `Filing set ${input.id} label is invalid`,
      );
    if (
      input.pageCap !== undefined &&
      (!Number.isInteger(input.pageCap) || input.pageCap < 1)
    )
      throw new AgentDocxError(
        "INVALID_ARGUMENT",
        `Filing set ${input.id} pageCap is invalid`,
      );
    if (!Array.isArray(input.documentIds) || input.documentIds.length === 0)
      throw new AgentDocxError(
        "INVALID_ARGUMENT",
        `Filing set ${input.id} documentIds must be a nonempty array`,
      );
    const existing = opened.manifest.filingSets ?? [];
    if (existing.some((entry) => entry.id === input.id))
      throw new AgentDocxError("PROJECT_INVALID", "Filing set already exists");
    const references: string[] = [];
    const seen = new Set<string>();
    for (const reference of input.documentIds) {
      if (typeof reference !== "string")
        throw new AgentDocxError(
          "INVALID_ARGUMENT",
          `Filing set ${input.id} document id is invalid`,
        );
      if (seen.has(reference))
        throw new AgentDocxError(
          "PROJECT_INVALID",
          `Filing set ${input.id} references duplicate document ${reference}`,
        );
      if (!opened.manifest.documents.some((entry) => entry.id === reference))
        throw new AgentDocxError(
          "PROJECT_INVALID",
          `Filing set ${input.id} references unknown document ${reference}`,
        );
      seen.add(reference);
      references.push(reference);
    }
    const filingSet: FilingSet = {
      id: input.id,
      ...(input.label !== undefined ? { label: input.label } : {}),
      documentIds: references,
      ...(input.pageCap !== undefined ? { pageCap: input.pageCap } : {}),
    };
    const filingSets = [...existing, filingSet].sort((left, right) =>
      left.id < right.id ? -1 : left.id > right.id ? 1 : 0,
    );
    const next = await updateManifest(opened, {
      ...opened.manifest,
      filingSets,
    });
    return getStateLocked(ctx, next);
  });
};

export const removeFilingSet = async (
  ctx: ProjectContext,
  id: string,
): Promise<ProjectState> => {
  return withLockedStore(ctx.manifestPath, async (opened) => {
    const existing = opened.manifest.filingSets ?? [];
    if (!existing.some((entry) => entry.id === id))
      throw new AgentDocxError("PROJECT_INVALID", "Filing set not found");
    const filingSets = existing.filter((entry) => entry.id !== id);
    const manifest: AgentDocxManifest = { ...opened.manifest };
    if (filingSets.length > 0) manifest.filingSets = filingSets;
    else delete manifest.filingSets;
    const next = await updateManifest(opened, manifest);
    return getStateLocked(ctx, next);
  });
};

export const getFilingSet = async (
  ctx: ProjectContext,
  id: string,
): Promise<FilingSetSnapshot> =>
  withLockedStore(ctx.manifestPath, async (opened) => {
    const filingSet = filingSetById(opened.manifest, id);
    const documents = await Promise.all(
      filingSet.documentIds.map(async (documentId) => {
        const config = documentById(opened.manifest, documentId);
        const snapshot = await snapshotProjectDocument(opened, config);
        const head = await readHead(opened.storePath, documentId);
        const record = head
          ? await readRevisionJson<RevisionRecord>(opened.storePath, head)
          : null;
        return {
          documentId,
          head,
          workingTreeHash: snapshot.workingTreeHash,
          matchesHead: record?.workingTreeHash === snapshot.workingTreeHash,
        };
      }),
    );
    return {
      schemaVersion: 1,
      id: filingSet.id,
      label: filingSet.label ?? null,
      documentIds: [...filingSet.documentIds],
      pageCap: filingSet.pageCap ?? null,
      documents,
    };
  });

export const validateFilingSet = async (
  ctx: ProjectContext,
  id: string,
): Promise<FilingSetValidation> => {
  return withLockedStore(ctx.manifestPath, async (opened) => {
    const filingSet = filingSetById(opened.manifest, id);
    const documents: FilingSetValidation["documents"][number][] = [];
    for (const documentId of filingSet.documentIds) {
      const head = await readHead(opened.storePath, documentId);
      let validation: ValidationResult | null = null;
      let pageCount: number | null = null;
      if (head !== null) {
        validation = await validateLocked(opened, documentId);
        pageCount = (await measureLocked(opened, documentId)).deterministic
          .pageCount;
      }
      documents.push({ documentId, head, validation, pageCount });
    }
    const pageCap =
      filingSet.pageCap === undefined
        ? null
        : (() => {
            const totalPages = documents.reduce(
              (total, document) => total + (document.pageCount ?? 0),
              0,
            );
            const missing = documents.find(
              (document) =>
                document.head === null || document.pageCount === null,
            );
            if (missing) {
              const missingValue =
                missing.head === null ? "no HEAD" : "no page count";
              return {
                limit: filingSet.pageCap!,
                totalPages,
                status: "unknown" as const,
                detail: `Document ${missing.documentId} has ${missingValue}; page cap cannot be evaluated`,
              };
            }
            return {
              limit: filingSet.pageCap!,
              totalPages,
              status: (totalPages <= filingSet.pageCap!
                ? "pass"
                : "fail") as FilingSetValidation["status"],
              detail:
                totalPages <= filingSet.pageCap!
                  ? `Total pages ${totalPages} are within page cap ${filingSet.pageCap!}`
                  : `Total pages ${totalPages} exceed page cap ${filingSet.pageCap!}`,
            };
          })();
    let status: FilingSetValidation["status"] = "pass";
    const rank: Record<FilingSetValidation["status"], number> = {
      pass: 0,
      unknown: 1,
      fail: 2,
    };
    for (const document of documents) {
      const candidate = (document.validation?.status ??
        "unknown") as FilingSetValidation["status"];
      if (rank[candidate] > rank[status]) status = candidate;
    }
    if (pageCap && rank[pageCap.status] > rank[status]) status = pageCap.status;
    return {
      schemaVersion: 1,
      id: filingSet.id,
      documents,
      pageCap,
      status,
    };
  });
};

const filingSetById = (
  manifest: AgentDocxManifest,
  filingSetId: string,
): FilingSet => {
  const filingSet = manifest.filingSets?.find(
    (entry) => entry.id === filingSetId,
  );
  if (!filingSet)
    throw new AgentDocxError("PROJECT_INVALID", "Filing set not found");
  return filingSet;
};
