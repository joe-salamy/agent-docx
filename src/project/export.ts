import { AgentDocxError } from "../types.js";
import {
  assertNoSymlinkComponents,
  canonicalJson,
  clearExportIntent,
  completeExportIntent,
  objectId,
  updateExportIntent,
  withLockedStore,
  writeObject,
  type ExportIntent,
  type OpenedStore,
} from "./store.js";
import {
  configuredRulePacks,
  sourceAssets,
  sourceFontSet,
} from "./documents.js";
import {
  currentRevision,
  isFirstParentAncestor,
  materialFor,
  snapshotForMaterial,
} from "./revisions.js";
import { sourcePathFor, version } from "./index.js";
import {
  createChangeSet,
  defaultAttribution,
} from "../revisions/diff.js";
import { compileMarkdown, createSemanticManifest } from "../docx/compile.js";
import { generateRedlineDocx } from "../docx/redline.js";
import { renderLibreOffice } from "../renderers/office.js";
import { dirname, extname, isAbsolute, relative, resolve, sep } from "node:path";
import { lstat, mkdir, rm } from "node:fs/promises";
import { inspectDocxMaterial } from "../docx/import.js";
import { pathExists, pathsOverlap, writeExclusiveFile } from "./fs-util.js";
import type { AgentDocxDocumentConfig } from "./contracts.js";
import type {
  ExportDocxInput,
  GeneratedAttachmentBundle,
  ProjectCompiledDocx,
} from "../docx/contracts.js";
import type { RevisionId } from "../legal/model.js";
import type { ProjectContext } from "./context.js";

export const exportDocx = async (ctx: ProjectContext, 
    documentId: string,
    input: ExportDocxInput,
  ): Promise<ProjectCompiledDocx> => {
    return withLockedStore(ctx.manifestPath, async (opened) => {
      const record = await currentRevision(ctx, 
        opened,
        documentId,
        input.revision,
      );
      const material = await materialFor(ctx, opened, record);
      const snapshot = await snapshotForMaterial(ctx, opened, material);
      const fontSet = sourceFontSet(material.config, snapshot);
      const generationDependencies = new Map(
        [...snapshot.dependencyBytes].map(([key, dependency]) => {
          const sha256 = record.dependencyObjects[key];
          if (!sha256)
            throw new AgentDocxError(
              "PROJECT_INVALID",
              `Missing stored dependency hash: ${key}`,
            );
          return [key, { ...dependency, sha256 }] as const;
        }),
      );
      const compiled = await compileMarkdown(
        material.source,
        {
          projectId: opened.manifest.projectId,
          documentId,
          profile: material.config.profile,
          ...(material.config.filingKind !== undefined
            ? { filingKind: material.config.filingKind }
            : {}),
          ...(material.config.rulePack !== undefined
            ? { rulePack: material.config.rulePack }
            : {}),
          metadata: material.config.metadata,
          ...(material.config.chrome !== undefined
            ? { chrome: material.config.chrome }
            : {}),
          ...(snapshot.dependencyBytes.get("template")
            ? { template: snapshot.dependencyBytes.get("template")!.bytes }
            : {}),
          ...(fontSet ? { fontSet } : {}),
          assets: sourceAssets(snapshot),
        },
        {
          ...input.options,
          ...(input.mode === "pdf"
            ? { renderer: "deterministic" as const }
            : {}),
          rulePacks: await configuredRulePacks(
            opened,
            material.config,
            snapshot,
          ),
          generation: {
            revision: record,
            annotations: input.mode === "redline" ? material.annotations : [],
            dependencies: generationDependencies,
          },
        },
      );
      let bytes = compiled.bytes;
      let mode: "clean" | "redline" | "pdf" = "clean";
      let baseRevision: RevisionId | null = null;
      let redlineVerification:
        | { revisionCount: number; commentCount: number; fieldCount: number }
        | undefined;
      if (input.mode === "redline") {
        const base = await currentRevision(ctx, 
          opened,
          documentId,
          input.baseRevision,
        );
        if (
          base.id === record.id ||
          !(await isFirstParentAncestor(ctx, opened, base.id, record))
        )
          throw new AgentDocxError(
            "DOCX_REDLINE_UNSUPPORTED",
            "Redline base must be a distinct first-parent ancestor",
          );
        if (base.documentConfigObject !== record.documentConfigObject)
          throw new AgentDocxError(
            "DOCX_REDLINE_UNSUPPORTED",
            "Redline export does not support document configuration changes",
          );
        const baseMaterial = await materialFor(ctx, opened, base);
        const changeSet = createChangeSet(
          documentId,
          base.id,
          record.id,
          baseMaterial.document,
          material.document,
          baseMaterial.annotations,
          material.annotations,
          defaultAttribution(record.author, record.createdAt),
          {
            baseSource: baseMaterial.source,
            headSource: material.source,
          },
        );
        try {
          const generated = await generateRedlineDocx(
            baseMaterial.document,
            material.document,
            changeSet,
            compiled.measurement.deterministic.profile,
            {
              ...(material.config.chrome !== undefined
                ? { chrome: material.config.chrome }
                : {}),
              metadata: material.config.metadata,
              pageCount: Math.max(
                1,
                compiled.measurement.deterministic.pageCount,
              ),
              semanticManifest: createSemanticManifest({
                document: material.document,
                source: material.source,
                mode: "redline",
                attachments: compiled.attachments?.manifest ?? null,
                revision: record.id,
                baseRevision: base.id,
                validation: compiled.validation,
                dependencies: generationDependencies,
                changeSet,
                annotations: material.annotations,
              }),
              createdAt: record.createdAt,
            },
          );
          bytes = generated.bytes;
          redlineVerification = {
            revisionCount: generated.revisionCount,
            commentCount: generated.commentCount,
            fieldCount: 0,
          };
        } catch (error) {
          throw new AgentDocxError(
            "DOCX_REDLINE_UNSUPPORTED",
            error instanceof Error ? error.message : String(error),
          );
        }
        mode = "redline";
        baseRevision = base.id;
      }

      const importedAttachments = compiled.attachments
        ? {
            files: compiled.attachments.files,
            manifest: compiled.attachments.manifest,
          }
        : undefined;
      const inspected = await inspectDocxMaterial(bytes, {
        ...(importedAttachments ? { attachments: importedAttachments } : {}),
      });
      if (
        inspected.result.fidelity.overall === "unsupported" ||
        !inspected.semantic ||
        inspected.semantic.projectId !== opened.manifest.projectId ||
        inspected.semantic.documentId !== documentId ||
        inspected.semantic.mode !== mode ||
        inspected.semantic.revision !== record.id ||
        inspected.semantic.baseRevision !== baseRevision
      )
        throw new AgentDocxError(
          "DOCX_IMPORT_UNSUPPORTED",
          "Generated DOCX failed strict semantic re-import validation",
        );
      const pdfRendering =
        input.mode === "pdf"
          ? await renderLibreOffice(
              bytes,
              [compiled.measurement.deterministic.profile.requestedFontFamily],
              {
                ...(input.options?.libreoffice ?? {}),
                includePdfBytes: true,
              },
              input.options?.officeTimeoutMs ?? 60000,
            )
          : null;
      if (pdfRendering) mode = "pdf";
      if (pdfRendering && !pdfRendering.pdf)
        throw new AgentDocxError(
          "LIBREOFFICE_RENDER_FAILED",
          "LibreOffice renderer did not return PDF bytes",
        );
      const pdfBytes = pdfRendering?.pdf ?? null;

      const destination = await assertExportDestination(
        opened,
        material.config,
        input.output,
      );
      const owner = ctx.randomUuid();
      const stagePath = `${destination.output}.agent-docx-${owner}.stage`;
      if (await pathExists(stagePath))
        throw new AgentDocxError(
          "OUTPUT_EXISTS",
          `DOCX export stage already exists: ${stagePath}`,
        );
      const emptyObject = objectId(new Uint8Array());
      const initialIntent: ExportIntent = {
        schemaVersion: 1,
        state: "preparing",
        projectId: opened.manifest.projectId,
        manifestPath: opened.manifestPath,
        owner,
        outputPath: destination.output,
        attachmentPath: null,
        stagePath,
        docxStagePath: resolve(stagePath, "document.docx"),
        pdfStagePath:
          input.mode === "pdf" ? resolve(stagePath, "document.pdf") : null,
        attachmentStagePath: null,
        docxSha256: emptyObject,
        attachmentManifestSha256: null,
        artifactProvenanceSha256: emptyObject,
        artifactStorePath: resolve(
          opened.storePath,
          "artifacts",
          "0".repeat(64),
          "0".repeat(64),
          "document.docx",
        ),
        pdfStorePath: null,
        pdfSha256: emptyObject,
        attachmentStorePath: null,
      };
      await updateExportIntent(opened.projectDirectory, initialIntent);
      let pdfObject: RevisionId | null = null;
      let pdfStorePath: string | null = null;
      let artifactObject!: RevisionId;
      let artifactProvenance!: RevisionId;
      let artifactStorePath!: string;
      let attachmentStorePath: string | null = null;
      let attachmentPath: string | null = null;
      let path = destination.output;
      let publicationStarted = false;
      try {
        await createExportStage(
          stagePath,
          owner,
          opened.manifest.projectId,
          opened.manifestPath,
        );
        await writeExclusiveFile(resolve(stagePath, "document.docx"), bytes);
        if (pdfBytes)
          await writeExclusiveFile(
            resolve(stagePath, "document.pdf"),
            pdfBytes,
          );
        if (compiled.attachments) {
          attachmentPath = destination.attachment;
          await writeAttachmentStage(stagePath, compiled.attachments);
        }
        artifactObject = await writeObject(opened.storePath, bytes);
        if (pdfBytes) pdfObject = await writeObject(opened.storePath, pdfBytes);
        if (compiled.attachments) {
          await writeObject(
            opened.storePath,
            canonicalJson(compiled.attachments.manifest),
          );
          for (const entry of compiled.attachments.manifest.entries)
            await writeObject(
              opened.storePath,
              compiled.attachments.files[entry.name]!.bytes,
            );
        }
        const provenanceJson = canonicalJson({
          schemaVersion: 1,
          generator: "agent-docx",
          generatorVersion: version,
          documentId,
          revision: record.id,
          mode,
          baseRevision,
          profile: compiled.artifact.profile,
          rulePack: compiled.artifact.rulePack,
          dependencies: Object.fromEntries(
            Object.entries(record.dependencyObjects).sort(([left], [right]) =>
              left.localeCompare(right),
            ),
          ),
          docxSha256: artifactObject,
          ...(pdfRendering && pdfObject
            ? {
                pdfSha256: pdfObject,
                pdfPageCount: pdfRendering.pageCount,
                pdfDelta:
                  pdfRendering.pageCount -
                  compiled.measurement.deterministic.pageCount,
                pdfRendererProvenance: {
                  versionRaw: pdfRendering.versionRaw,
                  executablePath: pdfRendering.executablePath,
                  platform: pdfRendering.platform,
                  arch: pdfRendering.arch,
                  calibratedFontEnvironment:
                    pdfRendering.calibratedFontEnvironment,
                  requestedFontFamilies: pdfRendering.requestedFontFamilies,
                  durationMs: pdfRendering.durationMs,
                  generatedDocxSha256: pdfRendering.generatedDocxSha256,
                  pdfSha256: pdfRendering.pdfSha256,
                },
              }
            : {}),
          attachments: compiled.attachments
            ? [...compiled.attachments.manifest.entries]
                .map((entry) => ({
                  name: entry.name,
                  mediaType: entry.mediaType,
                  byteLength: entry.byteLength,
                  sha256: entry.sha256,
                  payloadPath: entry.payloadPath,
                }))
                .sort((left, right) => left.name.localeCompare(right.name))
            : null,
          attachmentManifestSha256:
            compiled.attachments?.manifestSha256 ?? null,
        });
        artifactProvenance = objectId(provenanceJson);
        const artifactDirectory = artifactDirectoryFor(
          opened.storePath,
          record.id,
          artifactProvenance,
        );
        pdfStorePath = pdfBytes
          ? resolve(artifactDirectory, "document.pdf")
          : null;
        artifactStorePath = resolve(artifactDirectory, "document.docx");
        attachmentStorePath = compiled.attachments
          ? resolve(artifactDirectory, "attachments", "manifest.json")
          : null;
        const artifactStagePath = resolve(stagePath, "artifact");
        await mkdir(artifactStagePath, { mode: 0o700 });
        await writeExclusiveFile(
          resolve(artifactStagePath, "document.docx"),
          bytes,
        );
        if (pdfBytes)
          await writeExclusiveFile(
            resolve(artifactStagePath, "document.pdf"),
            pdfBytes,
          );
        await writeExclusiveFile(
          resolve(artifactStagePath, "provenance.json"),
          provenanceJson,
        );
        if (compiled.attachments)
          await writeAttachmentStage(artifactStagePath, compiled.attachments);
        await updateExportIntent(opened.projectDirectory, {
          ...initialIntent,
          state: "prepared",
          attachmentPath,
          attachmentStagePath: compiled.attachments
            ? resolve(stagePath, "attachments")
            : null,
          docxSha256: artifactObject,
          pdfStagePath: pdfBytes ? resolve(stagePath, "document.pdf") : null,
          pdfStorePath,
          pdfSha256: pdfObject ?? emptyObject,
          attachmentManifestSha256:
            compiled.attachments?.manifestSha256 ?? null,
          artifactProvenanceSha256: artifactProvenance,
          artifactStorePath,
          attachmentStorePath,
        });
        publicationStarted = true;
        await completeExportIntent(
          opened.projectDirectory,
          opened.manifestPath,
        );
      } catch (error) {
        if (!publicationStarted) {
          await rm(stagePath, { recursive: true, force: true }).catch(() => {});
          await clearExportIntent(opened.projectDirectory).catch(() => {});
        }
        throw error;
      }
      const artifact = {
        ...compiled.artifact,
        byteLength: bytes.byteLength,
        sha256: artifactObject,
        rendererProvenance: {
          ...compiled.artifact.rendererProvenance,
          ...(redlineVerification ? { verification: redlineVerification } : {}),
        },
        provenanceSha256: artifactProvenance,
        path,
        storePath: artifactStorePath,
        attachments:
          compiled.attachments && attachmentPath && attachmentStorePath
            ? {
                path: attachmentPath,
                storePath: attachmentStorePath,
                manifestSha256: compiled.attachments.manifestSha256,
                manifest: compiled.attachments.manifest,
              }
            : null,
        pdf:
          pdfRendering && pdfObject && pdfStorePath
            ? {
                sha256: pdfObject,
                pageCount: pdfRendering.pageCount,
                deterministicPageCount:
                  compiled.measurement.deterministic.pageCount,
                delta:
                  pdfRendering.pageCount -
                  compiled.measurement.deterministic.pageCount,
                rendererProvenance: {
                  versionRaw: pdfRendering.versionRaw,
                  executablePath: pdfRendering.executablePath,
                  platform: pdfRendering.platform,
                  arch: pdfRendering.arch,
                  calibratedFontEnvironment:
                    pdfRendering.calibratedFontEnvironment,
                  requestedFontFamilies: pdfRendering.requestedFontFamilies,
                  durationMs: pdfRendering.durationMs,
                  generatedDocxSha256: pdfRendering.generatedDocxSha256,
                  pdfSha256: pdfRendering.pdfSha256,
                },
                path,
                storePath: pdfStorePath,
              }
            : null,
        revision: record.id,
        mode,
        baseRevision,
      };
      return {
        ...compiled,
        bytes,
        measurement: {
          ...compiled.measurement,
          documentId,
          revision: record.id,
          workingTreeHash: record.workingTreeHash,
          dependencyObjects: record.dependencyObjects,
        },
        artifact,
      } as ProjectCompiledDocx;
    });
  };

const assertExportDestination = async (
  opened: OpenedStore,
  config: AgentDocxDocumentConfig,
  output: string,
): Promise<{ output: string; attachment: string }> => {
  const absoluteOutput = resolve(output);
  const relativeOutput = relative(opened.projectDirectory, absoluteOutput);
  if (
    relativeOutput === ".." ||
    relativeOutput.startsWith(`..${sep}`) ||
    isAbsolute(relativeOutput)
  )
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      "DOCX output must be inside the project directory",
    );
  await assertNoSymlinkComponents(absoluteOutput, "DOCX output");
  const attachment = attachmentDirectoryFor(absoluteOutput);
  const ownedPaths = exportOwnedPaths(opened, config);
  if (ownedPaths.some((owned) => pathsOverlap(absoluteOutput, owned))) {
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      "DOCX output overlaps a project source, dependency, or control path",
    );
  }
  await assertRegularDirectory(dirname(absoluteOutput), "DOCX output parent");
  if (await pathExists(absoluteOutput))
    throw new AgentDocxError(
      "OUTPUT_EXISTS",
      `DOCX output already exists: ${absoluteOutput}`,
    );
  if (await pathExists(attachment))
    throw new AgentDocxError(
      "OUTPUT_EXISTS",
      `Attachment bundle already exists: ${attachment}`,
    );
  return { output: absoluteOutput, attachment };
};


const exportOwnedPaths = (
  opened: OpenedStore,
  config: AgentDocxDocumentConfig,
): readonly string[] =>
  [
    opened.manifestPath,
    opened.storePath,
    sourcePathFor(opened, config),
    config.template ? resolve(opened.projectDirectory, config.template) : null,
    config.assetsDir
      ? resolve(opened.projectDirectory, config.assetsDir)
      : null,
    config.fontSet?.regularPath
      ? resolve(opened.projectDirectory, config.fontSet.regularPath)
      : null,
    config.fontSet?.boldPath
      ? resolve(opened.projectDirectory, config.fontSet.boldPath)
      : null,
    config.fontSet?.italicPath
      ? resolve(opened.projectDirectory, config.fontSet.italicPath)
      : null,
    config.fontSet?.boldItalicPath
      ? resolve(opened.projectDirectory, config.fontSet.boldItalicPath)
      : null,
  ].filter((path): path is string => path !== null);


const createExportStage = async (
  stagePath: string,
  owner: string,
  projectId: string,
  manifestPath: string,
): Promise<void> => {
  await mkdir(stagePath, { mode: 0o700 });
  await writeExclusiveFile(
    resolve(stagePath, "owner.json"),
    canonicalJson({ schemaVersion: 1, owner, projectId, manifestPath }),
  );
};


const writeAttachmentStage = async (
  stagePath: string,
  bundle: GeneratedAttachmentBundle,
): Promise<void> => {
  const attachmentRoot = resolve(stagePath, "attachments");
  await mkdir(attachmentRoot, { mode: 0o700 });
  for (const entry of bundle.manifest.entries) {
    if (
      !entry.payloadPath.startsWith("files/") ||
      entry.payloadPath.includes("\\")
    )
      throw new AgentDocxError(
        "PROJECT_INVALID",
        `Invalid attachment payload path: ${entry.payloadPath}`,
      );
    const payload = bundle.files[entry.name];
    if (!payload)
      throw new AgentDocxError(
        "PROJECT_INVALID",
        `Missing attachment bytes: ${entry.name}`,
      );
    const destination = resolve(attachmentRoot, entry.payloadPath);
    const relativePayload = relative(attachmentRoot, destination);
    if (
      relativePayload === ".." ||
      relativePayload.startsWith(`..${sep}`) ||
      isAbsolute(relativePayload)
    )
      throw new AgentDocxError(
        "PROJECT_INVALID",
        "Attachment payload escapes its bundle",
      );
    await mkdir(dirname(destination), { recursive: true, mode: 0o700 });
    await writeExclusiveFile(destination, payload.bytes);
  }
  await writeExclusiveFile(
    resolve(attachmentRoot, "manifest.json"),
    canonicalJson(bundle.manifest),
  );
};


const artifactDirectoryFor = (
  storePath: string,
  revision: RevisionId,
  provenance: RevisionId,
): string =>
  resolve(
    storePath,
    "artifacts",
    revision.slice("sha256:".length),
    provenance.slice("sha256:".length),
  );


const attachmentDirectoryFor = (output: string): string => {
  const path = resolve(output);
  const extension = extname(path);
  return `${extension.toLowerCase() === ".docx" ? path.slice(0, -extension.length) : path}.attachments`;
};


const assertRegularDirectory = async (
  path: string,
  label: string,
): Promise<void> => {
  let entry;
  try {
    entry = await lstat(path);
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === "ENOENT")
      throw new AgentDocxError(
        "INPUT_NOT_FOUND",
        `${label} does not exist: ${path}`,
      );
    throw error;
  }
  if (!entry.isDirectory() || entry.isSymbolicLink())
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `${label} is not a regular directory: ${path}`,
    );
};

