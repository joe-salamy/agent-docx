import { readFile, stat } from "node:fs/promises";
import { parseArgs } from "node:util";
import { dirname, isAbsolute, relative, resolve, sep } from "node:path";
import chokidar from "chokidar";
import type { CliCommand } from "./cli-args.js";
import {
  agentActions,
  dispatchAgentRequest,
  isStatelessAgentRequest,
  parseAgentRequest,
  serializeAgentValue,
  type AgentAction,
  type AgentRequest,
} from "./agent.js";
import { openProject } from "./project/index.js";
import { jsonlLines, MAX_JSONL_LINE_BYTES } from "./jsonl.js";
import {
  AgentDocxError,
  type CliErrorPayload,
  type JsonValue,
} from "./types.js";
import type { CliRuntime, CliSequenceState } from "./cli-run.js";

const string = { type: "string" } as const;
const boolean = { type: "boolean" } as const;

type Values = Record<string, string | boolean | undefined>;
type Parsed = { values: Values; positionals: readonly string[] };

const parse = (
  args: readonly string[],
  options: Record<string, { type: "string" | "boolean" }>,
): Parsed => {
  const parsed = parseArgs({
    args: [...args],
    options,
    strict: true,
    allowPositionals: true,
    tokens: true,
  });
  const seen = new Set<string>();
  for (const token of parsed.tokens) {
    if (token.kind !== "option" || !token.name) continue;
    if (seen.has(token.name))
      throw new AgentDocxError(
        "INVALID_ARGUMENT",
        `Duplicate option: --${token.name}`,
      );
    seen.add(token.name);
  }
  return { values: parsed.values as Values, positionals: parsed.positionals };
};

const required = (values: Values, key: string): string => {
  const value = values[key];
  if (typeof value !== "string" || value === "")
    throw new AgentDocxError("INVALID_ARGUMENT", `--${key} is required`);
  return value;
};

const optional = (values: Values, key: string): string | undefined =>
  typeof values[key] === "string" ? (values[key] as string) : undefined;

const noPositionals = (parsed: Parsed, label: string): void => {
  if (parsed.positionals.length > 0)
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `${label} does not accept positional arguments`,
    );
};

const readJson = async (
  cwd: string,
  path: string,
  label: string,
): Promise<unknown> => {
  try {
    return JSON.parse(await readFile(resolve(cwd, path), "utf8"));
  } catch (error) {
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `${label} must be valid JSON: ${error instanceof Error ? error.message : String(error)}`,
    );
  }
};

const actor = (values: Values): { name: string; email?: string } => {
  const name = required(values, "author");
  const email = optional(values, "email");
  return email === undefined ? { name } : { name, email };
};

const projectOption = (values: Values): string | undefined =>
  optional(values, "project");

const documentId = async (cwd: string, values: Values): Promise<string> => {
  const explicit = optional(values, "document");
  if (explicit) return explicit;
  const project = await openProject(
    resolve(cwd, projectOption(values) ?? "agent-docx.json"),
  );
  return (await project.getState()).manifest.defaultDocument;
};

const request = (
  action: AgentAction,
  project: string | undefined,
  params: Record<string, unknown>,
): AgentRequest => ({
  schemaVersion: 1,
  action,
  ...(project ? { project } : {}),
  params,
});
const print = async (
  runtime: CliRuntime,
  json: boolean,
  value: unknown,
): Promise<void> => {
  const serializable = serializeAgentValue(value, runtime.cwd);
  await runtime.writeStdout(
    `${JSON.stringify(serializable, null, json ? 0 : 2)}\n`,
  );
};

const errorPayload = (error: unknown): CliErrorPayload => {
  if (error instanceof AgentDocxError)
    return {
      code: error.code,
      message: error.message,
      ...(error.details ? { details: error.details } : {}),
    };
  return {
    code: "INTERNAL_ERROR",
    message: error instanceof Error ? error.message : String(error),
  };
};

const publicProjectPath = (cwd: string, path: string): string => {
  const project = relative(cwd, resolve(cwd, path)).split(sep).join("/");
  return project || ".";
};

const partialRequest = (value: unknown) => {
  const record =
    typeof value === "object" && value !== null && !Array.isArray(value)
      ? (value as Record<string, unknown>)
      : null;
  const action =
    record &&
    typeof record.action === "string" &&
    (agentActions as readonly string[]).includes(record.action)
      ? (record.action as AgentAction)
      : null;
  const params =
    record &&
    typeof record.params === "object" &&
    record.params !== null &&
    !Array.isArray(record.params)
      ? (record.params as Record<string, unknown>)
      : {};
  return {
    requestId:
      record &&
      (typeof record.id === "string" ||
        (typeof record.id === "number" && Number.isFinite(record.id)) ||
        record.id === null)
        ? record.id
        : null,
    action,
    project:
      record && typeof record.project === "string" && record.project !== ""
        ? record.project
        : undefined,
    stateless: action !== null && isStatelessAgentRequest(action, params),
  };
};

const parsePatchArgs = async (
  cwd: string,
  values: Values,
  targetDocument: string,
): Promise<unknown> => {
  const patch = optional(values, "patch");
  const base = optional(values, "base");
  const edits = optional(values, "edits");
  if (patch && (base || edits))
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      "Use --patch or --base with --edits, not both",
    );
  if (patch) return readJson(cwd, patch, "patch");
  if (!base || !edits)
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      "draft command requires --patch or both --base and --edits",
    );
  return {
    schemaVersion: 1,
    documentId: targetDocument,
    baseRevision: base,
    edits: await readJson(cwd, edits, "edits"),
  };
};

const commonProject = { project: string, json: boolean } as const;
const projectDocument = {
  ...commonProject,
  document: string,
  source: string,
  profile: string,
  metadata: string,
  chrome: string,
  template: string,
  "assets-dir": string,
  "filing-kind": string,
  "rule-pack": string,
  "create-source": boolean,
  "font-family": string,
  "font-regular": string,
  "font-bold": string,
  "font-italic": string,
  "font-bold-italic": string,
  default: boolean,
} as const;

const inputFromProjectFlags = async (cwd: string, values: Values) => {
  const metadata = await readJson(
    cwd,
    required(values, "metadata"),
    "metadata",
  );
  const chrome = optional(values, "chrome");
  const fontFamily = optional(values, "font-family");
  const fontRegular = optional(values, "font-regular");
  if ((fontFamily === undefined) !== (fontRegular === undefined))
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      "--font-family and --font-regular must be supplied together",
    );
  return {
    documentId: required(values, "document"),
    source: required(values, "source"),
    profile: required(values, "profile"),
    metadata,
    ...(values["create-source"] === true ? { createSource: true } : {}),
    ...(chrome ? { chrome: await readJson(cwd, chrome, "chrome") } : {}),
    ...(optional(values, "template")
      ? { template: optional(values, "template") }
      : {}),
    ...(optional(values, "assets-dir")
      ? { assetsDir: optional(values, "assets-dir") }
      : {}),
    ...(optional(values, "filing-kind")
      ? { filingKind: optional(values, "filing-kind") }
      : {}),
    ...(optional(values, "rule-pack")
      ? { rulePack: optional(values, "rule-pack") }
      : {}),
    ...(fontFamily && fontRegular
      ? {
          fontSet: {
            family: fontFamily,
            regularPath: fontRegular,
            ...(optional(values, "font-bold")
              ? { boldPath: optional(values, "font-bold") }
              : {}),
            ...(optional(values, "font-italic")
              ? { italicPath: optional(values, "font-italic") }
              : {}),
            ...(optional(values, "font-bold-italic")
              ? { boldItalicPath: optional(values, "font-bold-italic") }
              : {}),
          },
        }
      : {}),
  };
};

const nextAgentSequence = (state: CliSequenceState): number => ++state.sequence;

const runAgentInput = async (
  runtime: CliRuntime,
  project: string | undefined,
  sequence: CliSequenceState,
): Promise<number> => {
  if (runtime.stdinIsTTY)
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      "agent --input-jsonl requires non-TTY stdin",
    );
  for await (const line of jsonlLines(runtime)) {
    if (!line.trim()) continue;
    let raw: unknown = null;
    try {
      if (new TextEncoder().encode(line).byteLength > MAX_JSONL_LINE_BYTES)
        throw new AgentDocxError(
          "INVALID_ARGUMENT",
          `JSONL input line exceeds ${MAX_JSONL_LINE_BYTES} bytes`,
          { maxBytes: MAX_JSONL_LINE_BYTES },
        );
      raw = JSON.parse(line);
      const parsed = parseAgentRequest(raw);
      if (project && parsed.project !== undefined)
        throw new AgentDocxError(
          "INVALID_ARGUMENT",
          "Agent request project conflicts with command --project",
        );
      const result = await dispatchAgentRequest(
        project ? { ...parsed, project } : parsed,
        runtime.cwd,
      );
      await runtime.writeStdout(
        `${JSON.stringify({
          schemaVersion: 1,
          kind: "result",
          sequence: nextAgentSequence(sequence),
          requestId: result.request.id ?? null,
          action: result.request.action,
          project: result.project,
          documentId: result.documentId,
          revision: result.revision,
          value: serializeAgentValue(result.value, runtime.cwd),
        })}\n`,
      );
    } catch (error) {
      const partial = partialRequest(raw);
      const selectedProject =
        partial.action === null || partial.stateless
          ? null
          : publicProjectPath(
              runtime.cwd,
              project ?? partial.project ?? "agent-docx.json",
            );
      await runtime.writeStdout(
        `${JSON.stringify({
          schemaVersion: 1,
          kind: "error",
          sequence: nextAgentSequence(sequence),
          requestId: partial.requestId,
          action: partial.action,
          project: selectedProject,
          documentId: null,
          revision: null,
          error: errorPayload(error),
        })}\n`,
      );
    }
  }
  return 0;
};

type WatchDependencyState = {
  entries: readonly {
    key: string;
    path: string;
    byteLength: number;
    sha256: string;
  }[];
  paths: readonly string[];
};

const watchDependencies = async (
  cwd: string,
  projectPath: string,
  documentId: string,
): Promise<WatchDependencyState> => {
  const manifestPath = resolve(cwd, projectPath);
  const projectDirectory = dirname(manifestPath);
  const project = await openProject(manifestPath);
  const state = await project.getState();
  const target = state.documents.find(
    (document) => document.documentId === documentId,
  );
  const config = state.manifest.documents.find(
    (document) => document.id === documentId,
  );
  if (!target || !config)
    throw new AgentDocxError(
      "DOCUMENT_NOT_FOUND",
      `Document not found: ${documentId}`,
    );

  const entries: {
    key: string;
    path: string;
    byteLength: number;
    sha256: string;
  }[] = [];
  const paths = new Set<string>([manifestPath]);
  const add = async (
    key: string,
    path: string,
    sha256: string,
  ): Promise<void> => {
    paths.add(path);
    entries.push({
      key,
      path: relative(cwd, path).split(sep).join("/") || ".",
      byteLength: (await stat(path)).size,
      sha256,
    });
  };

  await add(
    "source",
    resolve(projectDirectory, config.source),
    target.sourceSha256,
  );
  if (config.template)
    await add(
      "template",
      resolve(projectDirectory, config.template),
      target.dependencyObjects.template!,
    );
  if (config.fontSet) {
    const fontPaths = {
      regular: config.fontSet.regularPath,
      bold: config.fontSet.boldPath,
      italic: config.fontSet.italicPath,
      boldItalic: config.fontSet.boldItalicPath,
    };
    for (const [role, path] of Object.entries(fontPaths)) {
      if (!path) continue;
      await add(
        `font/${role}`,
        resolve(projectDirectory, path),
        target.dependencyObjects[`font/${role}`]!,
      );
    }
  }
  if (config.assetsDir) {
    const assetsRoot = resolve(projectDirectory, config.assetsDir);
    paths.add(assetsRoot);
    for (const [key, sha256] of Object.entries(target.dependencyObjects)
      .filter(([key]) => key.startsWith("asset/"))
      .sort(([left], [right]) => left.localeCompare(right))) {
      const relativeAssetPath = key.slice("asset/".length);
      const assetPath = resolve(assetsRoot, relativeAssetPath);
      const assetRelative = relative(assetsRoot, assetPath);
      if (
        assetRelative === "" ||
        assetRelative === ".." ||
        assetRelative.startsWith(`..${sep}`) ||
        isAbsolute(assetRelative)
      )
        throw new AgentDocxError(
          "PROJECT_INVALID",
          `Invalid asset dependency: ${key}`,
        );
      await add(key, assetPath, sha256);
    }
  }
  return {
    entries: entries.sort((left, right) => left.key.localeCompare(right.key)),
    paths: [...paths].sort(),
  };
};

const runAgentWatch = async (
  runtime: CliRuntime,
  values: Values,
  sequence: CliSequenceState,
): Promise<number> => {
  if (values.jsonl !== true || values["input-jsonl"] === true)
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      "agent --watch requires --jsonl and forbids --input-jsonl",
    );
  const projectPath = required(values, "project");
  const targetDocument = required(values, "document");
  const publicProject = publicProjectPath(runtime.cwd, projectPath);
  const initialDependencies = await watchDependencies(
    runtime.cwd,
    projectPath,
    targetDocument,
  );
  const project = await openProject(resolve(runtime.cwd, projectPath));
  const state = await project.getState();
  const initial = state.documents.find(
    (document) => document.documentId === targetDocument,
  );
  if (!initial)
    throw new AgentDocxError(
      "DOCUMENT_NOT_FOUND",
      `Document not found: ${targetDocument}`,
    );

  let watchedPaths = new Set(initialDependencies.paths);
  const watcher = chokidar.watch([...watchedPaths], { ignoreInitial: true });
  await new Promise<void>((resolveReady, rejectReady) => {
    watcher.once("ready", resolveReady).once("error", rejectReady);
  });
  const emit = async (record: JsonValue): Promise<void> =>
    runtime.writeStdout(`${JSON.stringify(record)}\n`);
  await emit({
    schemaVersion: 1,
    kind: "ready",
    sequence: nextAgentSequence(sequence),
    requestId: null,
    action: "project.watch",
    project: publicProject,
    documentId: targetDocument,
    revision: initial.head,
    value: {
      protocolVersion: 1,
      capabilities: [...agentActions].sort(),
      dependencies: initialDependencies.entries,
    },
  });

  const refreshWatchedPaths = async (): Promise<void> => {
    const next = await watchDependencies(
      runtime.cwd,
      projectPath,
      targetDocument,
    );
    const nextPaths = new Set(next.paths);
    const removed = [...watchedPaths].filter((path) => !nextPaths.has(path));
    const added = [...nextPaths].filter((path) => !watchedPaths.has(path));
    if (removed.length) await watcher.unwatch(removed);
    if (added.length) watcher.add(added);
    watchedPaths = nextPaths;
  };

  let closing = false;
  let timer: NodeJS.Timeout | undefined;
  let pending = Promise.resolve();
  let lastRevision = initial.head;
  const emitError = async (error: unknown): Promise<void> => {
    if (closing) return;
    await emit({
      schemaVersion: 1,
      kind: "error",
      sequence: nextAgentSequence(sequence),
      requestId: null,
      action: "document.measure",
      project: publicProject,
      documentId: targetDocument,
      revision: null,
      error: errorPayload(error),
    });
  };
  const measure = async (): Promise<void> => {
    if (closing) return;
    try {
      const refreshed = await openProject(resolve(runtime.cwd, projectPath));
      const result = await refreshed.measure(targetDocument);
      await refreshWatchedPaths();
      if (closing) return;
      lastRevision = result.revision;
      await emit({
        schemaVersion: 1,
        kind: "result",
        sequence: nextAgentSequence(sequence),
        requestId: null,
        action: "document.measure",
        project: publicProject,
        documentId: targetDocument,
        revision: result.revision,
        value: serializeAgentValue(result, runtime.cwd),
      });
    } catch (error) {
      await emitError(error);
    }
  };
  const enqueue = (operation: () => Promise<void>): void => {
    pending = pending.then(operation);
    void pending.catch(() => undefined);
  };
  const queue = (): void => {
    if (closing) return;
    clearTimeout(timer);
    timer = setTimeout(() => {
      timer = undefined;
      enqueue(measure);
    }, 100);
  };
  watcher
    .on("change", queue)
    .on("add", queue)
    .on("unlink", queue)
    .on("addDir", queue)
    .on("unlinkDir", queue)
    .on("error", (error) => {
      if (!closing) enqueue(() => emitError(error));
    });

  const completion = Promise.withResolvers<number>();
  const stop = async (reason: "SIGINT" | "SIGTERM"): Promise<void> => {
    if (closing) return;
    closing = true;
    clearTimeout(timer);
    try {
      await pending;
      await watcher.close();
      await emit({
        schemaVersion: 1,
        kind: "end",
        sequence: nextAgentSequence(sequence),
        requestId: null,
        action: "project.watch",
        project: publicProject,
        documentId: targetDocument,
        revision: lastRevision,
        reason,
      });
      completion.resolve(0);
    } catch (error) {
      completion.reject(error);
    }
  };
  runtime.onceSignal("SIGINT", () => void stop("SIGINT"));
  runtime.onceSignal("SIGTERM", () => void stop("SIGTERM"));
  return completion.promise;
};

/** Runs every explicit non-measure CLI workflow command. */
export const runWorkflowCommand = async (
  command: Extract<CliCommand, { mode: "workflow" }>,
  runtime: CliRuntime,
  sequence: CliSequenceState,
): Promise<number> => {
  const [subcommand, ...rest] = command.args;
  let action: AgentAction;
  let params: Record<string, unknown>;
  let parsed: Parsed;
  switch (command.command) {
    case "project": {
      parsed = parse(rest, projectDocument);
      noPositionals(parsed, `project ${subcommand}`);
      params = await inputFromProjectFlags(runtime.cwd, parsed.values);
      if (subcommand === "add" && parsed.values.default === true)
        params.makeDefault = true;
      action = subcommand === "init" ? "project.init" : "project.add";
      break;
    }
    case "document": {
      parsed = parse(rest, {
        ...commonProject,
        document: string,
        base: string,
        changes: string,
        author: string,
        email: string,
        message: string,
      });
      noPositionals(parsed, "document configure");
      params = {
        documentId: await documentId(runtime.cwd, parsed.values),
        baseRevision: optional(parsed.values, "base") ?? null,
        changes: await readJson(
          runtime.cwd,
          required(parsed.values, "changes"),
          "changes",
        ),
        author: actor(parsed.values),
        message: required(parsed.values, "message"),
      };
      action = "document.configure";
      break;
    }
    case "revision": {
      if (subcommand === "checkpoint") {
        parsed = parse(rest, {
          ...commonProject,
          document: string,
          base: string,
          author: string,
          email: string,
          message: string,
        });
        noPositionals(parsed, "revision checkpoint");
        params = {
          documentId: await documentId(runtime.cwd, parsed.values),
          baseRevision: optional(parsed.values, "base") ?? null,
          author: actor(parsed.values),
          message: required(parsed.values, "message"),
        };
        action = "revision.checkpoint";
      } else if (subcommand === "list") {
        parsed = parse(rest, {
          ...commonProject,
          document: string,
          limit: string,
          cursor: string,
        });
        noPositionals(parsed, "revision list");
        const limit = optional(parsed.values, "limit");
        params = {
          documentId: await documentId(runtime.cwd, parsed.values),
          ...(limit ? { limit: Number(limit) } : {}),
          ...(optional(parsed.values, "cursor")
            ? { cursor: optional(parsed.values, "cursor") }
            : {}),
        };
        action = "revision.list";
      } else if (subcommand === "show") {
        parsed = parse(rest, {
          ...commonProject,
          document: string,
          revision: string,
        });
        if (parsed.positionals.length > 1)
          throw new AgentDocxError(
            "INVALID_ARGUMENT",
            "revision show accepts one revision",
          );
        params = {
          documentId: await documentId(runtime.cwd, parsed.values),
          revision:
            optional(parsed.values, "revision") ?? parsed.positionals[0],
        };
        action = "revision.get";
      } else if (subcommand === "restore") {
        parsed = parse(rest, {
          ...commonProject,
          document: string,
          base: string,
          author: string,
          email: string,
          message: string,
        });
        if (parsed.positionals.length !== 1)
          throw new AgentDocxError(
            "INVALID_ARGUMENT",
            "revision restore requires a target revision positional",
          );
        params = {
          documentId: await documentId(runtime.cwd, parsed.values),
          baseRevision: required(parsed.values, "base"),
          targetRevision: parsed.positionals[0],
          author: actor(parsed.values),
          message: required(parsed.values, "message"),
        };
        action = "revision.restore";
      } else if (subcommand === "diff") {
        parsed = parse(rest, {
          ...commonProject,
          document: string,
          base: string,
          head: string,
          output: string,
        });
        if (parsed.positionals.length > 2)
          throw new AgentDocxError(
            "INVALID_ARGUMENT",
            "revision diff accepts BASE and HEAD positionals",
          );
        params = {
          documentId: await documentId(runtime.cwd, parsed.values),
          baseRevision:
            optional(parsed.values, "base") ?? parsed.positionals[0],
          headRevision:
            optional(parsed.values, "head") ?? parsed.positionals[1],
          ...(optional(parsed.values, "output")
            ? { output: optional(parsed.values, "output") }
            : {}),
        };
        action = "revision.diff";
      } else {
        parsed = parse(rest, {
          ...commonProject,
          document: string,
          "change-set": string,
          decisions: string,
          author: string,
          email: string,
          message: string,
        });
        noPositionals(parsed, "revision resolve");
        params = {
          documentId: await documentId(runtime.cwd, parsed.values),
          changeSet: await readJson(
            runtime.cwd,
            required(parsed.values, "change-set"),
            "change set",
          ),
          decisions: await readJson(
            runtime.cwd,
            required(parsed.values, "decisions"),
            "decisions",
          ),
          author: actor(parsed.values),
          message: required(parsed.values, "message"),
        };
        action = "revision.resolve";
      }
      break;
    }
    case "draft": {
      if (subcommand === "guidance") {
        parsed = parse(rest, {
          ...commonProject,
          document: string,
          revision: string,
        });
        noPositionals(parsed, "draft guidance");
        params = {
          documentId: await documentId(runtime.cwd, parsed.values),
          ...(optional(parsed.values, "revision")
            ? { revision: optional(parsed.values, "revision") }
            : {}),
        };
        action = "draft.guidance";
      } else {
        parsed = parse(rest, {
          ...commonProject,
          document: string,
          patch: string,
          base: string,
          edits: string,
          renderer: string,
          "patch-hash": string,
          gate: string,
          author: string,
          email: string,
          message: string,
        });
        noPositionals(parsed, `draft ${subcommand}`);
        const targetDocument = await documentId(runtime.cwd, parsed.values);
        params = {
          patch: await parsePatchArgs(
            runtime.cwd,
            parsed.values,
            targetDocument,
          ),
          ...(optional(parsed.values, "renderer")
            ? { renderer: optional(parsed.values, "renderer") }
            : {}),
        };
        if (subcommand === "apply") {
          params.patchHash = required(parsed.values, "patch-hash");
          if (optional(parsed.values, "gate"))
            params.gate = optional(parsed.values, "gate");
          params.author = actor(parsed.values);
          params.message = required(parsed.values, "message");
          delete params.renderer;
          action = "draft.apply";
        } else action = "draft.evaluate";
      }
      break;
    }
    case "review": {
      if (subcommand === "add") {
        parsed = parse(rest, {
          ...commonProject,
          document: string,
          revision: string,
          block: string,
          start: string,
          end: string,
          author: string,
          email: string,
          message: string,
        });
        noPositionals(parsed, "review add");
        const start = optional(parsed.values, "start");
        const end = optional(parsed.values, "end");
        if ((start === undefined) !== (end === undefined))
          throw new AgentDocxError(
            "INVALID_ARGUMENT",
            "--start and --end must be supplied together",
          );
        params = {
          documentId: await documentId(runtime.cwd, parsed.values),
          revision: required(parsed.values, "revision"),
          blockId: required(parsed.values, "block"),
          ...(start && end
            ? { range: { start: Number(start), end: Number(end) } }
            : {}),
          author: actor(parsed.values),
          message: required(parsed.values, "message"),
        };
        action = "review.add";
      } else {
        parsed = parse(rest, {
          ...commonProject,
          document: string,
          revision: string,
          annotation: string,
          author: string,
          email: string,
          message: string,
        });
        noPositionals(parsed, "review resolve");
        params = {
          documentId: await documentId(runtime.cwd, parsed.values),
          revision: required(parsed.values, "revision"),
          annotationId: required(parsed.values, "annotation"),
          author: actor(parsed.values),
          message: required(parsed.values, "message"),
        };
        action = "review.resolve";
      }
      break;
    }
    case "validate": {
      parsed = parse(command.args, {
        ...commonProject,
        document: string,
        revision: string,
      });
      noPositionals(parsed, "validate");
      params = {
        documentId: await documentId(runtime.cwd, parsed.values),
        ...(optional(parsed.values, "revision")
          ? { revision: optional(parsed.values, "revision") }
          : {}),
      };
      action = "document.validate";
      break;
    }
    case "export": {
      parsed = parse(command.args, {
        ...commonProject,
        document: string,
        revision: string,
        mode: string,
        base: string,
        output: string,
        renderer: string,
        "office-timeout": string,
        "libreoffice-path": string,
      });
      noPositionals(parsed, "export");
      const mode = required(parsed.values, "mode");
      const renderer = optional(parsed.values, "renderer");
      const officeTimeoutMs = optional(parsed.values, "office-timeout");
      const libreofficePath = optional(parsed.values, "libreoffice-path");
      params = {
        documentId: await documentId(runtime.cwd, parsed.values),
        revision: required(parsed.values, "revision"),
        mode,
        output: required(parsed.values, "output"),
        ...(optional(parsed.values, "base")
          ? { baseRevision: optional(parsed.values, "base") }
          : {}),
        ...(renderer || officeTimeoutMs || libreofficePath
          ? {
              options: {
                ...(renderer ? { renderer } : {}),
                ...(officeTimeoutMs
                  ? { officeTimeoutMs: Number(officeTimeoutMs) }
                  : {}),
                ...(libreofficePath
                  ? { libreoffice: { executablePath: libreofficePath } }
                  : {}),
              },
            }
          : {}),
      };
      action = "docx.export";
      break;
    }
    case "import": {
      parsed = parse(command.args, {
        ...commonProject,
        document: string,
        output: string,
        attachments: string,
        "inspect-only": boolean,
        author: string,
        email: string,
        message: string,
      });
      if (parsed.positionals.length !== 1)
        throw new AgentDocxError(
          "INVALID_ARGUMENT",
          "import requires one DOCX input positional",
        );
      const inspectOnly = parsed.values["inspect-only"] === true;
      if (inspectOnly) {
        for (const key of [
          "project",
          "document",
          "output",
          "author",
          "email",
          "message",
        ] as const)
          if (parsed.values[key] !== undefined)
            throw new AgentDocxError(
              "INVALID_ARGUMENT",
              "import --inspect-only forbids project, document, output, and actor options",
            );
      }
      params = {
        input: parsed.positionals[0],
        inspectOnly,
        ...(optional(parsed.values, "attachments")
          ? { attachments: optional(parsed.values, "attachments") }
          : {}),
      };
      if (!inspectOnly)
        Object.assign(params, {
          documentId: await documentId(runtime.cwd, parsed.values),
          output: required(parsed.values, "output"),
          author: actor(parsed.values),
          message: required(parsed.values, "message"),
        });
      action = "docx.import";
      break;
    }
    case "import-redline": {
      parsed = parse(command.args, {
        ...commonProject,
        document: string,
        input: string,
        attachments: string,
        author: string,
        email: string,
        message: string,
      });
      noPositionals(parsed, "import-redline");
      params = {
        documentId: await documentId(runtime.cwd, parsed.values),
        input: required(parsed.values, "input"),
        author: actor(parsed.values),
        message: required(parsed.values, "message"),
        ...(optional(parsed.values, "attachments")
          ? { attachments: optional(parsed.values, "attachments") }
          : {}),
      };
      action = "docx.importRedline";
      break;
    }
    case "filing-set": {
      if (subcommand === "add") {
        parsed = parse(rest, {
          ...commonProject,
          id: string,
          label: string,
          documents: string,
          "page-cap": string,
        });
        noPositionals(parsed, "filing-set add");
        const documents = required(parsed.values, "documents")
          .split(",")
          .map((entry) => entry.trim())
          .filter((entry) => entry.length > 0);
        if (documents.length === 0)
          throw new AgentDocxError(
            "INVALID_ARGUMENT",
            "--documents requires a comma-separated document ID list",
          );
        params = {
          id: required(parsed.values, "id"),
          documentIds: documents,
          ...(optional(parsed.values, "label")
            ? { label: optional(parsed.values, "label") }
            : {}),
          ...(optional(parsed.values, "page-cap")
            ? { pageCap: Number(optional(parsed.values, "page-cap")) }
            : {}),
        };
        action = "filingSet.add";
      } else {
        parsed = parse(rest, { ...commonProject, id: string });
        noPositionals(parsed, `filing-set ${subcommand}`);
        params = { id: required(parsed.values, "id") };
        action =
          subcommand === "remove"
            ? "filingSet.remove"
            : subcommand === "get"
              ? "filingSet.get"
              : "filingSet.validate";
      }
      break;
    }
    case "agent": {
      parsed = parse(command.args, {
        project: string,
        document: string,
        "input-jsonl": boolean,
        watch: boolean,
        jsonl: boolean,
      });
      noPositionals(parsed, "agent");
      if (parsed.values.watch === true)
        return runAgentWatch(runtime, parsed.values, sequence);
      if (parsed.values["input-jsonl"] !== true || parsed.values.jsonl === true)
        throw new AgentDocxError(
          "INVALID_ARGUMENT",
          "agent requires --input-jsonl, or --watch --project --document --jsonl",
        );
      return runAgentInput(runtime, projectOption(parsed.values), sequence);
    }
  }
  const result = await dispatchAgentRequest(
    request(action, projectOption(parsed!.values), params!),
    runtime.cwd,
  );
  await print(runtime, parsed!.values.json === true, result.value);
  return 0;
};
