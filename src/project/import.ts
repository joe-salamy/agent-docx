import { AgentDocxError } from "../types.js";
import {
  assertNoSymlinkComponents,
  canonicalJson,
  readHead,
  replaceOwnedFile,
  snapshotProjectDocument,
  withLockedStore,
} from "./store.js";
import { lstat, mkdir, rm } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { commitLocked, importedAssetDestinations } from "./documents.js";
import {
  currentRevision,
  isFirstParentAncestor,
  materialFor,
  provenanceForRevision,
} from "./revisions.js";
import { documentById, sourcePathFor } from "./index.js";
import { parseLegalMarkdown } from "../legal/parse.js";
import { semanticDocumentProjection } from "../docx/compile.js";
import { inspectDocxMaterial, inspectRedlineResolution } from "../docx/import.js";
import {
  createChangeSet,
  defaultAttribution,
  reattributeChangeSet,
  type JsonObject,
} from "../revisions/diff.js";
import { writeExclusiveFile } from "./fs-util.js";
import type {
  DocxImportResult,
  ImportAttachmentBundle,
  ImportDocxInput,
  RedlineImportResult,
} from "../docx/contracts.js";
import type { Actor } from "../legal/model.js";
import type { ProjectSnapshot } from "./store.js";
import type { ProjectContext } from "./context.js";

export const importDocx = async (ctx: ProjectContext, 
    input: Extract<ImportDocxInput, { inspectOnly: false }>,
  ): Promise<DocxImportResult> => {
    const inspected = await inspectDocxMaterial(input.input, {
      ...(input.attachments ? { attachments: input.attachments } : {}),
    });
    if (inspected.result.fidelity.overall === "unsupported")
      throw new AgentDocxError(
        "DOCX_IMPORT_UNSUPPORTED",
        "DOCX contains constructs that cannot be faithfully imported",
      );
    return withLockedStore(ctx.manifestPath, async (opened) => {
      const config = documentById(opened.manifest, input.documentId);
      if (!inspected.semantic)
        throw new AgentDocxError(
          "DOCX_IMPORT_UNSUPPORTED",
          "Strict DOCX import requires an agent-docx semantic manifest",
        );
      if (
        inspected.semantic.projectId !== opened.manifest.projectId ||
        inspected.semantic.documentId !== config.id
      )
        throw new AgentDocxError(
          "DOCX_IMPORT_UNSUPPORTED",
          "DOCX semantic manifest belongs to a different project document",
        );
      const sourcePath = sourcePathFor(opened, config);
      if (resolve(input.output) !== sourcePath)
        throw new AgentDocxError(
          "PROJECT_INVALID",
          "Import output must be the configured project document source",
        );
      if ((await readHead(opened.storePath, input.documentId)) !== null)
        throw new AgentDocxError(
          "REVISION_CONFLICT",
          "Strict DOCX import requires an empty document history",
        );
      const snapshot = await snapshotProjectDocument(opened, config);
      if (snapshot.source.length !== 0)
        throw new AgentDocxError(
          "WORKING_COPY_CONFLICT",
          "Strict DOCX import requires a zero-byte configured source",
        );
      const assetDestinations = importedAssetDestinations(
        opened,
        config,
        inspected.assets,
      );
      for (const asset of assetDestinations) {
        await assertNoSymlinkComponents(asset.path, "Imported asset");
        try {
          await lstat(asset.path);
        } catch (error) {
          if ((error as NodeJS.ErrnoException).code === "ENOENT") continue;
          throw error;
        }
        throw new AgentDocxError(
          "WORKING_COPY_CONFLICT",
          `Strict DOCX import requires an absent asset target: ${asset.source}`,
        );
      }
      if (inspected.tracked) {
        const baseDocument = parseLegalMarkdown(inspected.tracked.baseSource, {
          projectId: opened.manifest.projectId,
          documentId: config.id,
          metadata: config.metadata,
          ...(config.chrome !== undefined
            ? { chrome: config.chrome }
            : {}),
          assets: inspected.assets,
          requireMarkers: true,
        }).document;
        const headDocument = parseLegalMarkdown(inspected.tracked.headSource, {
          projectId: opened.manifest.projectId,
          documentId: config.id,
          metadata: config.metadata,
          ...(config.chrome !== undefined
            ? { chrome: config.chrome }
            : {}),
          assets: inspected.assets,
          requireMarkers: true,
          annotations: inspected.result.recognized.annotations,
        }).document;
        if (
          canonicalJson(semanticDocumentProjection(headDocument)) !==
          canonicalJson(inspected.semantic.document)
        )
          throw new AgentDocxError(
            "DOCX_IMPORT_UNSUPPORTED",
            "DOCX redline head document does not match the configured project document",
          );
        const writtenAssets: string[] = [];
        let currentSnapshot = snapshot;
        let baseCommitted = false;
        try {
          for (const asset of assetDestinations) {
            await mkdir(dirname(asset.path), { recursive: true, mode: 0o700 });
            await writeExclusiveFile(asset.path, asset.bytes);
            writtenAssets.push(asset.path);
          }
          await replaceOwnedFile(
            sourcePath,
            currentSnapshot.sourceObject,
            inspected.tracked.baseSource,
          );
          currentSnapshot = await snapshotProjectDocument(opened, config);
          const base = await commitLocked(ctx, 
            opened,
            config,
            currentSnapshot,
            baseDocument,
            [],
            null,
            input.author,
            input.message,
            undefined,
            false,
            true,
            { expectedWorkingTreeHash: currentSnapshot.workingTreeHash },
          );
          baseCommitted = true;
          await replaceOwnedFile(
            sourcePath,
            currentSnapshot.sourceObject,
            inspected.tracked.headSource,
          );
          currentSnapshot = await snapshotProjectDocument(opened, config);
          const head = await commitLocked(ctx, 
            opened,
            config,
            currentSnapshot,
            headDocument,
            inspected.result.recognized.annotations,
            base.revision.id,
            input.author,
            input.message,
            undefined,
            false,
            true,
            { expectedWorkingTreeHash: currentSnapshot.workingTreeHash },
          );
          return {
            ...inspected.result,
            inspectOnly: false,
            mode: "tracked",
            output: sourcePath,
            sourceSha256: currentSnapshot.sourceObject,
            baseRevision: base.revision.id,
            headRevision: head.revision.id,
            revisions: [base.revision.id, head.revision.id],
            recognized: {
              ...inspected.result.recognized,
              blocks: headDocument.blocks,
              footnotes: headDocument.footnotes,
            },
          } as DocxImportResult;
        } catch (error) {
          if (currentSnapshot.source !== snapshot.source)
            await replaceOwnedFile(
              sourcePath,
              currentSnapshot.sourceObject,
              snapshot.source,
            );
          if (baseCommitted)
            await rm(resolve(opened.storePath, "refs", `${config.id}.json`), {
              force: true,
            });
          for (const path of writtenAssets.reverse())
            await rm(path, { force: true });
          throw error;
        }
      }
      const document = parseLegalMarkdown(inspected.source, {
        projectId: opened.manifest.projectId,
        documentId: config.id,
        metadata: config.metadata,
        ...(config.chrome !== undefined
          ? { chrome: config.chrome }
          : {}),
        assets: inspected.assets,
        requireMarkers: true,
        annotations: inspected.result.recognized.annotations,
      }).document;
      if (
        canonicalJson(semanticDocumentProjection(document)) !==
        canonicalJson(inspected.semantic.document)
      )
        throw new AgentDocxError(
          "DOCX_IMPORT_UNSUPPORTED",
          "DOCX semantic document does not match the configured project document",
        );
      let published = false;
      let refreshed: ProjectSnapshot | null = null;
      const writtenAssets: string[] = [];
      try {
        for (const asset of assetDestinations) {
          await mkdir(dirname(asset.path), { recursive: true, mode: 0o700 });
          await writeExclusiveFile(asset.path, asset.bytes);
          writtenAssets.push(asset.path);
        }
        await replaceOwnedFile(
          sourcePath,
          snapshot.sourceObject,
          inspected.source,
        );
        published = true;
        refreshed = await snapshotProjectDocument(opened, config);
        const mutation = await commitLocked(ctx, 
          opened,
          config,
          refreshed,
          document,
          inspected.result.recognized.annotations,
          null,
          input.author,
          input.message,
          undefined,
          false,
          true,
          { expectedWorkingTreeHash: refreshed.workingTreeHash },
        );
        return {
          ...inspected.result,
          inspectOnly: false,
          mode: "clean",
          output: sourcePath,
          sourceSha256: refreshed.sourceObject,
          baseRevision: mutation.revision.id,
          headRevision: mutation.revision.id,
          revisions: [mutation.revision.id],
          recognized: {
            ...inspected.result.recognized,
            blocks: document.blocks,
            footnotes: document.footnotes,
          },
        } as DocxImportResult;
      } catch (error) {
        if (published && refreshed)
          await replaceOwnedFile(
            sourcePath,
            refreshed.sourceObject,
            snapshot.source,
          );
        for (const path of writtenAssets.reverse())
          await rm(path, { force: true });
        throw error;
      }
    });
  };

export const importRedline = async (ctx: ProjectContext, input: {
    documentId: string;
    input: string | Uint8Array;
    attachments?: ImportAttachmentBundle;
    author: Actor;
    message: string;
  }): Promise<RedlineImportResult> => {
    const inspected = await inspectRedlineResolution(input.input, {
      ...(input.attachments ? { attachments: input.attachments } : {}),
    });
    return withLockedStore(ctx.manifestPath, async (opened) => {
      const config = documentById(opened.manifest, input.documentId);
      if (
        inspected.semantic.projectId !== opened.manifest.projectId ||
        inspected.semantic.documentId !== config.id
      )
        throw new AgentDocxError(
          "DOCX_IMPORT_UNSUPPORTED",
          "DOCX semantic manifest belongs to a different project document",
        );
      const semanticRevision = inspected.semantic.revision;
      const semanticBaseRevision = inspected.semantic.baseRevision;
      if (!semanticRevision || !semanticBaseRevision)
        throw new AgentDocxError(
          "REVISION_NOT_FOUND",
          "Redline semantic manifest does not identify committed revisions",
        );
      const currentHeadId = await readHead(opened.storePath, config.id);
      if (currentHeadId === null)
        throw new AgentDocxError(
          "REVISION_NOT_FOUND",
          `Document has no revision: ${config.id}`,
        );
      if (currentHeadId !== semanticRevision)
        throw new AgentDocxError(
          "REVISION_CONFLICT",
          "Redline semantic revision must be the current document head",
        );
      const head = await currentRevision(ctx, 
        opened,
        config.id,
        semanticRevision,
      );
      const base = await currentRevision(ctx, 
        opened,
        config.id,
        semanticBaseRevision,
      );
      if (
        base.id === head.id ||
        !(await isFirstParentAncestor(ctx, opened, base.id, head))
      )
        throw new AgentDocxError(
          "REVISION_CONFLICT",
          "Redline base must be a distinct first-parent ancestor of the head",
        );
      const baseMaterial = await materialFor(ctx, opened, base);
      const headMaterial = await materialFor(ctx, opened, head);
      if (inspected.semantic.source !== headMaterial.source)
        throw new AgentDocxError(
          "DOCX_IMPORT_UNSUPPORTED",
          "Redline semantic source does not match the committed head revision",
        );
      if (
        canonicalJson(semanticDocumentProjection(headMaterial.document)) !==
        canonicalJson(inspected.semantic.document)
      )
        throw new AgentDocxError(
          "DOCX_IMPORT_UNSUPPORTED",
          "Redline semantic document does not match the committed head revision",
        );
      const snapshot = await snapshotProjectDocument(opened, config);
      if (
        snapshot.workingTreeHash !== head.workingTreeHash ||
        snapshot.sourceObject !== head.sourceObject ||
        snapshot.documentConfigObject !== head.documentConfigObject ||
        canonicalJson(snapshot.dependencyObjects) !==
          canonicalJson(head.dependencyObjects)
      )
        throw new AgentDocxError(
          "WORKING_COPY_CONFLICT",
          "Working copy differs from the redline head revision",
        );
      if (
        inspected.resolution === "complete" &&
        inspected.rejectedSource === null
      )
        throw new AgentDocxError(
          "CHANGESET_INVALID",
          "Complete redline resolution did not produce a rejected source",
        );
      const rawChangeSet = createChangeSet(
        config.id,
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
      const changeSet = reattributeChangeSet(rawChangeSet, {
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
      const decisions: Record<`c_${string}`, "accept" | "reject"> = {};
      if (inspected.resolution === "complete") {
        const rawIds = new Set(rawChangeSet.changes.map((change) => change.id));
        for (const changeId of Object.keys(inspected.decisions))
          if (!rawIds.has(changeId as `c_${string}`))
            throw new AgentDocxError(
              "CHANGESET_INVALID",
              `Redline decision does not identify a committed change: ${changeId}`,
            );
        for (const [index, change] of changeSet.changes.entries()) {
          const rawChange = rawChangeSet.changes[index];
          const decision =
            rawChange === undefined
              ? undefined
              : inspected.decisions[rawChange.id];
          if (!decision)
            throw new AgentDocxError(
              "CHANGESET_INVALID",
              `Change cannot be attributed to a reviewer decision: ${change.id}`,
            );
          decisions[change.id] = decision;
        }
        const reviewerAnnotations = new Map(
          inspected.annotations.map((annotation) => [
            annotation.id,
            annotation,
          ]),
        );
        for (const change of changeSet.annotations) {
          if (change.kind === "add") {
            decisions[change.id] = reviewerAnnotations.has(change.newValue.id)
              ? "accept"
              : "reject";
          } else if (change.kind === "remove") {
            decisions[change.id] = reviewerAnnotations.has(change.oldValue.id)
              ? "reject"
              : "accept";
          } else {
            const current = reviewerAnnotations.get(change.newValue.id);
            if (
              current &&
              canonicalJson(current) === canonicalJson(change.newValue)
            )
              decisions[change.id] = "accept";
            else if (
              reviewerAnnotations.get(change.oldValue.id) &&
              canonicalJson(reviewerAnnotations.get(change.oldValue.id)) ===
                canonicalJson(change.oldValue)
            )
              decisions[change.id] = "reject";
            else
              throw new AgentDocxError(
                "CHANGESET_INVALID",
                `Annotation change cannot be attributed to a reviewer decision: ${change.id}`,
              );
          }
        }
      }
      const fidelity: RedlineImportResult["fidelity"] = {
        overall: "normalized",
        items: [
          {
            status: "preserved",
            partPath: "customXml/itemAgentDocx.xml",
            relationshipId: "rIdAgentDocxSemantic",
            ooxmlKind: "agent-docx:semantic-manifest",
            count: 1,
            blockIds: [],
            sourcePositions: [],
            explanation:
              "The agent-docx semantic manifest preserves revision and block identity.",
          },
          {
            status: "normalized",
            partPath: "word/document.xml",
            relationshipId: null,
            ooxmlKind: "w:ins/w:del",
            count: inspected.tracked.paragraphs.length,
            blockIds: inspected.tracked.paragraphs.flatMap((paragraph) =>
              paragraph.bookmark ? [paragraph.bookmark] : [],
            ),
            sourcePositions: [],
            explanation:
              "Reviewer redline markup was validated and normalized to semantic decisions.",
          },
        ],
      };
      return {
        schemaVersion: 1,
        documentId: config.id,
        baseRevision: base.id,
        headRevision: head.id,
        changeSet,
        decisions,
        resolution: inspected.resolution,
        annotations: inspected.annotations,
        fidelity,
      };
    });
  };
