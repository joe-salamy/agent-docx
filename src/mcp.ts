import { readFile, realpath } from "node:fs/promises";
import {
  basename,
  dirname,
  isAbsolute,
  relative,
  resolve,
  sep,
} from "node:path";
import {
  agentActions,
  executeAgentRequest,
  serializeAgentValue,
} from "./agent.js";
import { toErrorPayload } from "./errors.js";
import { MAX_JSONL_LINE_BYTES } from "./jsonl.js";
import type { CliRuntime } from "./cli-contract.js";
import { AgentDocxError } from "./types.js";

export type McpRuntime = Pick<
  CliRuntime,
  "cwd" | "version" | "readStdinChunks" | "writeStdout" | "onceSignal"
>;

type RpcId = string | number | null;
type RpcError = {
  code: number;
  message: string;
  data?: Record<string, unknown>;
};
type RpcReply = {
  jsonrpc: "2.0";
  id: RpcId;
  result?: unknown;
  error?: RpcError;
};

type PlainRecord = Record<string, unknown>;

const SUPPORTED_PROTOCOL_VERSIONS = [
  "2024-11-05",
  "2025-03-26",
  "2025-06-18",
] as const;

const hasOwn = (value: PlainRecord, key: string): boolean =>
  Object.prototype.hasOwnProperty.call(value, key);

const isRecord = (value: unknown): value is PlainRecord =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const isRpcId = (value: unknown): value is RpcId =>
  value === null ||
  typeof value === "string" ||
  (typeof value === "number" && Number.isFinite(value));

const isPlainObject = (value: unknown): value is PlainRecord =>
  isRecord(value) && Object.getPrototypeOf(value) === Object.prototype;

const rpcError = (
  id: RpcId,
  code: number,
  message: string,
  data?: Record<string, unknown>,
): RpcReply => ({
  jsonrpc: "2.0",
  id,
  error: {
    code,
    message,
    ...(data === undefined ? {} : { data }),
  },
});

const rpcResult = (id: RpcId, result: unknown): RpcReply => ({
  jsonrpc: "2.0",
  id,
  result,
});

const errorMessage = (error: unknown): string => {
  if (error instanceof Error) return error.message;
  if (isRecord(error) && typeof error.message === "string")
    return error.message;
  return String(error);
};

const isEpipe = (error: unknown): boolean =>
  isRecord(error) && error.code === "EPIPE";

type ActionSchema = {
  properties?: Record<string, unknown>;
  required?: readonly unknown[];
};

type LoadedActionSchemas = {
  byAction: Map<string, ActionSchema>;
  defs: Record<string, unknown>;
};

type SchemaBranch = {
  $ref?: string;
  properties?: PlainRecord;
  required?: readonly unknown[];
};

const resolveActionDef = (
  name: string,
  defs: Record<string, unknown>,
): ActionSchema => {
  const def = defs[name];
  if (typeof def !== "object" || def === null || Array.isArray(def)) return {};
  const { properties, required, allOf } = def as ActionSchema & {
    allOf?: readonly SchemaBranch[];
  };
  const merged: ActionSchema = {
    ...(properties ? { properties } : {}),
    ...(required ? { required } : {}),
  };
  for (const branch of allOf ?? []) {
    const inner =
      typeof branch.$ref === "string"
        ? resolveActionDef(branch.$ref.replace(/^#\/\$defs\//, ""), defs)
        : {
            ...(branch.properties ? { properties: branch.properties } : {}),
            ...(branch.required ? { required: branch.required } : {}),
          };
    merged.properties = {
      ...(inner.properties ?? {}),
      ...(merged.properties ?? {}),
    };
    merged.required = [...(inner.required ?? []), ...(merged.required ?? [])];
  }
  return merged;
};

const EXTERNAL_ACTION_SCHEMAS: ReadonlyArray<{
  id: string;
  file: string;
  defKey: string;
}> = [
  {
    id: "https://agent-docx.dev/schemas/change-set-v1.json",
    file: "../change-set.schema.json",
    defKey: "changeSet",
  },
  {
    id: "https://agent-docx.dev/schemas/source-patch-v1.json",
    file: "../source-patch.schema.json",
    defKey: "sourcePatch",
  },
];

const rewriteRefs = (
  node: unknown,
  rewrite: (value: string) => string,
): void => {
  if (Array.isArray(node)) {
    for (const item of node) rewriteRefs(item, rewrite);
    return;
  }
  if (typeof node !== "object" || node === null) return;
  for (const [key, value] of Object.entries(node)) {
    if (key === "$ref" && typeof value === "string") {
      (node as Record<string, unknown>)[key] = rewrite(value);
    } else {
      rewriteRefs(value, rewrite);
    }
  }
};

const loadActionSchemas = async (): Promise<LoadedActionSchemas | null> => {
  try {
    const raw = JSON.parse(
      await readFile(
        new URL("../agent-request.schema.json", import.meta.url),
        "utf8",
      ),
    ) as {
      $defs?: Record<string, unknown>;
      allOf?: Array<{
        oneOf?: Array<SchemaBranch & { properties?: PlainRecord }>;
      }>;
    };
    const defs: Record<string, unknown> = { ...(raw.$defs ?? {}) };
    for (const external of EXTERNAL_ACTION_SCHEMAS) {
      const bundled = JSON.parse(
        await readFile(new URL(external.file, import.meta.url), "utf8"),
      ) as Record<string, unknown>;
      delete bundled.$id;
      delete bundled.$schema;
      delete bundled.title;
      // Rewrite this bundle's own local refs so they resolve inside the
      // bundle once it is nested under the tool schema's $defs.
      const prefix = `#/$defs/${external.defKey}`;
      rewriteRefs(bundled, (value) =>
        value.startsWith("#/$defs/")
          ? `${prefix}/$defs/${value.slice("#/$defs/".length)}`
          : value,
      );
      defs[external.defKey] = bundled;
    }
    // Global pass: only the absolute external schema IDs become local refs.
    for (const external of EXTERNAL_ACTION_SCHEMAS) {
      const prefix = `#/$defs/${external.defKey}`;
      rewriteRefs(defs, (value) =>
        value.startsWith(external.id)
          ? `${prefix}${value.slice(external.id.length)}`
          : value,
      );
    }
    const byAction = new Map<string, ActionSchema>();
    for (const branch of raw.allOf?.[0]?.oneOf ?? []) {
      const action = branch.properties?.action as
        | { const?: string }
        | undefined;
      const params = branch.properties?.params as { $ref?: string } | undefined;
      if (typeof action?.const !== "string" || typeof params?.$ref !== "string")
        continue;
      const resolved = resolveActionDef(
        params.$ref.replace(/^#\/\$defs\//, ""),
        defs,
      );
      byAction.set(action.const, resolved);
    }
    return byAction.size > 0 ? { byAction, defs } : null;
  } catch {
    // Fall back to the generic schema; the file only loads under normal packaging.
    return null;
  }
};

const tools = (loaded: LoadedActionSchemas | null): PlainRecord[] =>
  agentActions.map((name) => {
    const params = loaded?.byAction.get(name);
    return {
      name,
      description:
        "agent-docx version-1 protocol action (params are validated as in agent-request.schema.json); project defaults to ./agent-docx.json",
      inputSchema: {
        type: "object",
        properties: {
          ...(params?.properties ?? {}),
          project: {
            type: "string",
            description: "Project manifest path, relative to the server cwd",
          },
        },
        ...(params?.required && params.required.length > 0
          ? { required: [...params.required] }
          : {}),
        ...(loaded ? { $defs: loaded.defs } : {}),
        additionalProperties: false,
      },
    };
  });

const initializeResult = (params: unknown, version: string): PlainRecord => {
  const requested = isRecord(params) ? params.protocolVersion : undefined;
  const protocolVersion =
    typeof requested === "string" &&
    (SUPPORTED_PROTOCOL_VERSIONS as readonly string[]).includes(requested)
      ? requested
      : "2025-06-18";
  return {
    protocolVersion,
    capabilities: { tools: { listChanged: false } },
    serverInfo: { name: "agent-docx", version },
  };
};

const canonicalizePath = async (candidate: string): Promise<string> => {
  let cursor = candidate;
  const suffix: string[] = [];
  for (;;) {
    try {
      const canonical = await realpath(cursor);
      return suffix.length
        ? resolve(canonical, ...suffix.reverse())
        : canonical;
    } catch (error) {
      const code =
        error !== null && typeof error === "object" && "code" in error
          ? error.code
          : undefined;
      if (code !== "ENOENT" && code !== "ENOTDIR") throw error;
      const parent = dirname(cursor);
      if (parent === cursor) throw error;
      suffix.push(basename(cursor));
      cursor = parent;
    }
  }
};

const confinedProjectPath = async (
  cwd: string,
  requested: string,
): Promise<string> => {
  let root: string;
  try {
    root = await realpath(cwd);
  } catch {
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      "project must stay inside the server cwd",
    );
  }
  const candidate = isAbsolute(requested) ? requested : resolve(cwd, requested);
  let canonical: string;
  try {
    canonical = await canonicalizePath(candidate);
  } catch {
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      "project must stay inside the server cwd",
    );
  }
  const escaped = relative(root, canonical);
  if (
    escaped !== "" &&
    (escaped === ".." || escaped.startsWith(`..${sep}`) || isAbsolute(escaped))
  )
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      "project must stay inside the server cwd",
    );
  return canonical;
};

const toolCallResult = async (
  id: RpcId,
  params: unknown,
  runtime: McpRuntime,
): Promise<RpcReply> => {
  if (!isRecord(params))
    return rpcError(id, -32602, "tools/call params must be an object");
  const name = params.name;
  if (typeof name !== "string" || name === "")
    return rpcError(id, -32602, "Tool name must be a non-empty string");
  if (!(agentActions as readonly string[]).includes(name))
    return rpcError(id, -32602, `Unknown tool: ${name}`);

  const rawArguments = params.arguments;
  if (rawArguments !== undefined && !isRecord(rawArguments))
    return rpcError(id, -32602, "Tool arguments must be an object");
  const argumentsObject = rawArguments === undefined ? {} : rawArguments;
  if (
    argumentsObject.project !== undefined &&
    (typeof argumentsObject.project !== "string" ||
      argumentsObject.project === "")
  )
    return rpcError(id, -32602, "project must be a non-empty string");
  const callParams: PlainRecord = { ...argumentsObject };
  delete callParams.project;

  try {
    const envelope: PlainRecord = {
      schemaVersion: 1,
      id,
      action: name,
      params: callParams,
    };
    if (typeof argumentsObject.project === "string")
      envelope.project = await confinedProjectPath(
        runtime.cwd,
        argumentsObject.project,
      );
    const result = await executeAgentRequest(envelope, runtime.cwd);
    const serialized = serializeAgentValue(result.value, runtime.cwd);
    const toolResult: PlainRecord = {
      content: [{ type: "text", text: JSON.stringify(serialized) }],
      isError: false,
    };
    if (isPlainObject(serialized)) toolResult.structuredContent = serialized;
    return rpcResult(id, toolResult);
  } catch (error) {
    const projected = toErrorPayload(error);
    return rpcResult(id, {
      content: [{ type: "text", text: projected.message }],
      structuredContent: {
        code: projected.code,
        message: projected.message,
        ...(projected.details ? { details: projected.details } : {}),
      },
      isError: true,
    });
  }
};

const handleMessage = async (
  line: string,
  runtime: McpRuntime,
  writeReply: (reply: RpcReply) => Promise<void>,
  actionSchemas: LoadedActionSchemas | null,
): Promise<void> => {
  let message: unknown;
  try {
    message = JSON.parse(line) as unknown;
  } catch (error) {
    await writeReply(
      rpcError(null, -32700, `Invalid JSON: ${errorMessage(error)}`),
    );
    return;
  }

  if (Array.isArray(message)) {
    await writeReply(
      rpcError(null, -32600, "Batch requests are not supported"),
    );
    return;
  }
  if (!isRecord(message)) {
    await writeReply(rpcError(null, -32600, "Invalid Request"));
    return;
  }

  const hasId = hasOwn(message, "id");
  const id = hasId && isRpcId(message.id) ? message.id : null;
  if (
    message.jsonrpc !== "2.0" ||
    typeof message.method !== "string" ||
    (hasId && !isRpcId(message.id))
  ) {
    await writeReply(rpcError(id, -32600, "Invalid Request"));
    return;
  }

  const method = message.method;
  if (!hasId) {
    if (method === "notifications/initialized") return;
    return;
  }

  switch (method) {
    case "initialize":
      await writeReply(
        rpcResult(id, initializeResult(message.params, runtime.version)),
      );
      return;
    case "ping":
      await writeReply(rpcResult(id, {}));
      return;
    case "tools/list":
      await writeReply(rpcResult(id, { tools: tools(actionSchemas) }));
      return;
    case "tools/call":
      await writeReply(await toolCallResult(id, message.params, runtime));
      return;
    default:
      await writeReply(rpcError(id, -32601, `Method not found: ${method}`));
  }
};

const waitForGrace = async (pending: Promise<unknown>): Promise<void> => {
  await Promise.race([
    pending.catch(() => undefined),
    new Promise<void>((resolve) => setTimeout(resolve, 100)),
  ]);
};

const closeIterator = async (
  iterator: AsyncIterator<Uint8Array>,
): Promise<void> => {
  if (typeof iterator.return !== "function") return;
  try {
    await iterator.return();
  } catch {
    // Stopping the transport must not turn a signal into a failure.
  }
};

export const runMcpServer = async (runtime: McpRuntime): Promise<number> => {
  const actionSchemas = await loadActionSchemas();
  let stopping = false;
  let resolveStop!: () => void;
  const stopPromise = new Promise<void>((resolve) => {
    resolveStop = resolve;
  });
  const requestStop = (): void => {
    if (stopping) return;
    stopping = true;
    resolveStop();
  };
  runtime.onceSignal("SIGINT", requestStop);
  runtime.onceSignal("SIGTERM", requestStop);

  let outputClosed = false;
  const writeReply = async (reply: RpcReply): Promise<void> => {
    if (outputClosed || stopping) return;
    try {
      await runtime.writeStdout(`${JSON.stringify(reply)}\n`);
    } catch (error) {
      if (!isEpipe(error)) throw error;
      outputClosed = true;
      requestStop();
    }
  };

  let iterator: AsyncIterator<Uint8Array>;
  try {
    iterator = runtime.readStdinChunks()[Symbol.asyncIterator]();
  } catch {
    return 1;
  }

  const decoder = new TextDecoder("utf-8", { fatal: true });
  let pending = "";
  const appendLines = (text: string): string[] => {
    const lines: string[] = [];
    let start = 0;
    while (start <= text.length) {
      const newline = text.indexOf("\n", start);
      const fragment = text.slice(
        start,
        newline === -1 ? text.length : newline,
      );
      const rawBytes = Buffer.byteLength(pending) + Buffer.byteLength(fragment);
      const normalizedBytes =
        newline !== -1 && fragment.endsWith("\r") ? rawBytes - 1 : rawBytes;
      if (normalizedBytes > MAX_JSONL_LINE_BYTES)
        throw new AgentDocxError(
          "INVALID_ARGUMENT",
          `JSON-RPC line exceeds ${MAX_JSONL_LINE_BYTES} bytes`,
        );
      if (newline === -1) {
        pending += fragment;
        break;
      }
      const rawLine = pending + fragment;
      pending = "";
      lines.push(rawLine.endsWith("\r") ? rawLine.slice(0, -1) : rawLine);
      start = newline + 1;
      if (start === text.length) break;
    }
    return lines;
  };

  const processLines = async (lines: readonly string[]): Promise<void> => {
    for (const line of lines) {
      if (stopping) return;
      if (line.trim() === "") continue;
      await handleMessage(line, runtime, writeReply, actionSchemas);
    }
  };
  const protocolError = async (error: unknown): Promise<void> => {
    await writeReply(
      rpcError(
        null,
        -32700,
        error instanceof AgentDocxError ? error.message : "Invalid UTF-8 input",
      ),
    );
  };

  while (!stopping) {
    const next = iterator.next().then(
      (result) => ({ kind: "chunk" as const, result }),
      (error: unknown) => ({ kind: "read-error" as const, error }),
    );
    const outcome = await Promise.race([
      next,
      stopPromise.then(() => ({ kind: "signal" as const })),
    ]);
    if (outcome.kind === "signal") {
      void closeIterator(iterator);
      await waitForGrace(next);
      return 0;
    }
    if (outcome.kind === "read-error") {
      void closeIterator(iterator);
      return 1;
    }
    if (outcome.result.done) break;

    let lines: string[];
    try {
      lines = appendLines(
        decoder.decode(outcome.result.value, { stream: true }),
      );
    } catch (error) {
      await protocolError(error);
      void closeIterator(iterator);
      return 0;
    }
    const processTask = processLines(lines);
    const processed = await Promise.race([
      processTask.then(
        () => ({ kind: "processed" as const }),
        (error: unknown) => ({ kind: "process-error" as const, error }),
      ),
      stopPromise.then(() => ({ kind: "signal" as const })),
    ]);
    if (processed.kind === "signal") {
      void closeIterator(iterator);
      await waitForGrace(processTask);
      return 0;
    }
    if (processed.kind === "process-error") {
      void closeIterator(iterator);
      return 1;
    }
  }

  if (stopping) {
    void closeIterator(iterator);
    return 0;
  }

  try {
    const lines = appendLines(decoder.decode());
    await processLines(lines);
    if (pending) {
      const line = pending;
      pending = "";
      await processLines([line]);
    }
  } catch (error) {
    await protocolError(error);
    return 0;
  }
  return 0;
};
