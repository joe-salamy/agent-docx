import { readFile } from "node:fs/promises";
import { inspectDocxTemplate } from "./docx/inspect.js";
import { inspectDocx } from "./docx/import.js";
import type { ExportDocxInput, ImportDocxInput } from "./docx/contracts.js";
import type {
  AddReviewInput,
  ConfigureDocumentInput,
  ProjectMeasureOptions,
  ResolveChangesInput,
  ResolveReviewInput,
} from "./project/contracts.js";
import { createProject, openProject } from "./project/index.js";
import type {
  AgentDispatchResult,
  AgentRequest,
} from "./agent-protocol.js";
import { responseMeta } from "./agent-serialize.js";
import {
  asChangeSet,
  asPatch,
  configUpdate,
  invocationPath,
  isStatelessAgentRequest,
  projectInput,
  projectPath,
  publicPath,
} from "./agent-protocol.js";

/** Executes one already-validated local agent request without exposing binary DOCX bytes. */
export const dispatchAgentRequest = async (
  request: AgentRequest,
  cwd = process.cwd(),
): Promise<AgentDispatchResult> => {
  const manifestPath = projectPath(cwd, request.project);
  const projectDisplay = isStatelessAgentRequest(request.action, request.params)
    ? null
    : publicPath(cwd, manifestPath);
  let value: unknown;
  switch (request.action) {
    case "project.init": {
      value = await (
        await createProject(
          manifestPath,
          projectInput(cwd, manifestPath, request.params, false),
        )
      ).getState();
      break;
    }
    case "project.add": {
      const project = await openProject(manifestPath);
      value = await project.addDocument(
        projectInput(cwd, manifestPath, request.params, true),
      );
      break;
    }
    case "project.get": {
      value = await (await openProject(manifestPath)).getState();
      break;
    }
    case "document.configure": {
      const input: ConfigureDocumentInput = {
        baseRevision: request.params.baseRevision,
        changes: configUpdate(
          cwd,
          manifestPath,
          request.params.changes,
        ),
        author: request.params.author,
        message: request.params.message,
      };
      value = await (
        await openProject(manifestPath)
      ).configureDocument(request.params.documentId, input);
      break;
    }
    case "document.get":
      value = await (
        await openProject(manifestPath)
      ).getDocument(request.params.documentId, request.params.revision);
      break;
    case "document.measure":
      value = await (
        await openProject(manifestPath)
      ).measure(
        request.params.documentId,
        request.params.revision,
        request.params.options as ProjectMeasureOptions | undefined,
      );
      break;
    case "document.validate":
      value = await (
        await openProject(manifestPath)
      ).validate(request.params.documentId, request.params.revision);
      break;
    case "revision.checkpoint": {
      value = await (
        await openProject(manifestPath)
      ).checkpoint(request.params.documentId, {
        baseRevision: request.params.baseRevision,
        author: request.params.author,
        message: request.params.message,
      });
      break;
    }
    case "revision.list": {
      value = await (
        await openProject(manifestPath)
      ).listRevisions(request.params.documentId, {
        ...(request.params.limit === undefined
          ? {}
          : { limit: request.params.limit }),
        ...(request.params.cursor === undefined
          ? {}
          : { cursor: request.params.cursor }),
      });
      break;
    }
    case "revision.get":
      value = await (
        await openProject(manifestPath)
      ).getRevision(request.params.documentId, request.params.revision);
      break;
    case "revision.restore":
      value = await (
        await openProject(manifestPath)
      ).restore(request.params.documentId, {
        baseRevision: request.params.baseRevision,
        targetRevision: request.params.targetRevision,
        author: request.params.author,
        message: request.params.message,
      });
      break;
    case "revision.diff": {
      const project = await openProject(manifestPath);
      const changeSet = await project.diff(
        request.params.documentId,
        request.params.baseRevision,
        request.params.headRevision,
      );
      value = request.params.output
        ? {
            changeSet,
            compiled: await project.exportDocx(request.params.documentId, {
              revision: request.params.headRevision,
              mode: "redline",
              baseRevision: request.params.baseRevision,
              output: invocationPath(cwd, request.params.output),
            }),
          }
        : { changeSet };
      break;
    }
    case "revision.resolve": {
      const input: ResolveChangesInput = {
        changeSet: asChangeSet(request.params.changeSet),
        decisions: request.params.decisions,
        author: request.params.author,
        message: request.params.message,
      };
      value = await (
        await openProject(manifestPath)
      ).resolveChanges(request.params.documentId, input);
      break;
    }
    case "draft.guidance":
      value = await (
        await openProject(manifestPath)
      ).getDraftGuidance(request.params.documentId, request.params.revision);
      break;
    case "draft.evaluate":
      value = await (
        await openProject(manifestPath)
      ).evaluatePatch(asPatch(request.params.patch), {
        ...(request.params.renderer === undefined
          ? {}
          : { renderer: request.params.renderer }),
      });
      break;
    case "draft.apply":
      value = await (
        await openProject(manifestPath)
      ).applyPatch(asPatch(request.params.patch), {
        patchHash: request.params.patchHash,
        ...(request.params.gate === undefined
          ? {}
          : { gate: request.params.gate }),
        author: request.params.author,
        message: request.params.message,
      });
      break;
    case "review.add": {
      const input: AddReviewInput = {
        revision: request.params.revision,
        blockId: request.params.blockId as AddReviewInput["blockId"],
        ...(request.params.range === undefined
          ? {}
          : { range: request.params.range }),
        author: request.params.author,
        message: request.params.message,
      };
      value = await (
        await openProject(manifestPath)
      ).addReview(request.params.documentId, input);
      break;
    }
    case "review.resolve": {
      const input: ResolveReviewInput = {
        revision: request.params.revision,
        annotationId: request.params.annotationId as ResolveReviewInput["annotationId"],
        author: request.params.author,
        message: request.params.message,
      };
      value = await (
        await openProject(manifestPath)
      ).resolveReview(request.params.documentId, input);
      break;
    }
    case "docx.export": {
      const output = invocationPath(cwd, request.params.output);
      const options = request.params.options as ExportDocxInput["options"];
      let input: ExportDocxInput;
      if (request.params.mode === "clean" || request.params.mode === "pdf") {
        input = {
          revision: request.params.revision,
          mode: request.params.mode,
          output,
          ...(options === undefined ? {} : { options }),
        };
      } else {
        input = {
          revision: request.params.revision,
          mode: "redline",
          baseRevision: request.params.baseRevision!,
          output,
          ...(options === undefined ? {} : { options }),
        };
      }
      value = await (
        await openProject(manifestPath)
      ).exportDocx(request.params.documentId, input);
      break;
    }
    case "docx.import": {
      const inputPath = invocationPath(cwd, request.params.input);
      const attachments =
        request.params.attachments === undefined
          ? undefined
          : { directory: invocationPath(cwd, request.params.attachments) };
      if (request.params.inspectOnly) {
        value = await inspectDocx(
          inputPath,
          attachments ? { attachments } : {},
        );
      } else {
        const input: Extract<ImportDocxInput, { inspectOnly: false }> = {
          input: inputPath,
          inspectOnly: false,
          documentId: request.params.documentId!,
          output: invocationPath(cwd, request.params.output!),
          author: request.params.author!,
          message: request.params.message!,
          ...(attachments ? { attachments } : {}),
        };
        value = await (await openProject(manifestPath)).importDocx(input);
      }
      break;
    }
    case "docx.inspect":
      value = await inspectDocxTemplate(
        await readFile(invocationPath(cwd, request.params.input)),
      );
      break;
    case "docx.importRedline": {
      const project = await openProject(manifestPath);
      value = await project.importRedline({
        documentId: request.params.documentId,
        input: invocationPath(cwd, request.params.input),
        ...(request.params.attachments === undefined
          ? {}
          : {
              attachments: {
                directory: invocationPath(cwd, request.params.attachments),
              },
            }),
        author: request.params.author,
        message: request.params.message,
      });
      break;
    }
    case "filingSet.add": {
      const project = await openProject(manifestPath);
      value = await project.addFilingSet({
        id: request.params.id,
        documentIds: request.params.documentIds,
        ...(request.params.label === undefined
          ? {}
          : { label: request.params.label }),
        ...(request.params.pageCap === undefined
          ? {}
          : { pageCap: request.params.pageCap }),
      });
      break;
    }
    case "filingSet.remove":
      value = await (
        await openProject(manifestPath)
      ).removeFilingSet(request.params.id);
      break;
    case "filingSet.get":
      value = await (
        await openProject(manifestPath)
      ).getFilingSet(request.params.id);
      break;
    case "filingSet.validate":
      value = await (
        await openProject(manifestPath)
      ).validateFilingSet(request.params.id);
      break;
  }
  const meta = responseMeta(request.action, request.params, value);
  return { request, project: projectDisplay, ...meta, value };
};
