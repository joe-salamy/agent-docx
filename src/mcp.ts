import {
  agentActions,
  dispatchAgentRequest,
  serializeAgentValue,
} from "./agent.js";
import { AgentDocxError, type ErrorCode } from "./types.js";

export type McpRuntime = {
  cwd: string;
  stdinIsTTY: boolean;
  readStdinChunks(): AsyncGenerator<Uint8Array>;
  writeStdout(text: string): Promise<void>;
  writeStderr(text: string): Promise<void>;
  onceSignal(signal: string, listener: () => void): void;
  version: string;
};

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

const errorCode = (error: unknown): ErrorCode | undefined =>
  error instanceof AgentDocxError ? error.code : undefined;

const errorMessage = (error: unknown): string => {
  if (error instanceof Error) return error.message;
  if (isRecord(error) && typeof error.message === "string")
    return error.message;
  return String(error);
};

const isEpipe = (error: unknown): boolean =>
  isRecord(error) && error.code === "EPIPE";

const tools = (): PlainRecord[] =>
  agentActions.map((name) => ({
    name,
    description:
      "agent-docx version-1 protocol action (params are validated as in agent-request.schema.json); project defaults to ./agent-docx.json",
    inputSchema: {
      type: "object",
      properties: {
        project: {
          type: "string",
          description: "Project manifest path, relative to the server cwd",
        },
      },
      additionalProperties: true,
    },
  }));

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
  const callParams: PlainRecord = { ...argumentsObject };
  delete callParams.project;

  const envelope: PlainRecord = {
    schemaVersion: 1,
    id,
    action: name,
    params: callParams,
  };
  if (typeof argumentsObject.project === "string")
    envelope.project = argumentsObject.project;

  try {
    const result = await dispatchAgentRequest(envelope, runtime.cwd);
    const serialized = serializeAgentValue(result.value, runtime.cwd);
    const toolResult: PlainRecord = {
      content: [{ type: "text", text: JSON.stringify(serialized) }],
      isError: false,
    };
    if (isPlainObject(serialized)) toolResult.structuredContent = serialized;
    return rpcResult(id, toolResult);
  } catch (error) {
    const code = errorCode(error);
    const message = errorMessage(error);
    return rpcResult(id, {
      content: [{ type: "text", text: message }],
      structuredContent: {
        code: code ?? "INTERNAL_ERROR",
        message,
      },
      isError: true,
    });
  }
};

const handleMessage = async (
  line: string,
  runtime: McpRuntime,
  writeReply: (reply: RpcReply) => Promise<void>,
): Promise<void> => {
  let message: unknown;
  try {
    message = JSON.parse(line) as unknown;
  } catch (error) {
    await writeReply(
      rpcError(null, -32602, `Invalid JSON: ${errorMessage(error)}`),
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
      await writeReply(rpcResult(id, { tools: tools() }));
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

  const decoder = new TextDecoder();
  let pending = "";

  const processLines = async (): Promise<void> => {
    while (!stopping) {
      const newline = pending.indexOf("\n");
      if (newline < 0) return;
      const line = pending.slice(0, newline);
      pending = pending.slice(newline + 1);
      if (line.trim() === "") continue;
      await handleMessage(line, runtime, writeReply);
    }
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

    pending += decoder.decode(outcome.result.value, { stream: true });
    const processTask = processLines();
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

  pending += decoder.decode();
  if (pending.trim() !== "") {
    const line = pending;
    pending = "";
    try {
      await handleMessage(line, runtime, writeReply);
    } catch {
      return 1;
    }
  }
  return 0;
};
