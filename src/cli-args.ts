import { parseArgs, type ParseArgsConfig } from "node:util";
import { AgentDocxError } from "./types.js";

export type ParseArgsToken =
  | {
      kind: "option";
      index: number;
      name: string;
      rawName: string;
      value: string | undefined;
      inlineValue: boolean | undefined;
    }
  | { kind: "positional"; index: number; value: string }
  | { kind: "option-terminator"; index: number };

export type StrictParsedArgs = {
  values: Record<string, string | boolean | readonly string[] | undefined>;
  positionals: string[];
  tokens?: readonly ParseArgsToken[];
};

/**
 * Strict parseArgs with a stable error taxonomy: unknown or malformed options
 * surface as `INVALID_ARGUMENT` rather than a raw node error.
 */
export const parseCliArgsStrict = (
  config: ParseArgsConfig & { args: readonly string[] },
): StrictParsedArgs => {
  try {
    const parsed = parseArgs({
      ...config,
      tokens: true,
    }) as StrictParsedArgs;
    const seen = new Set<string>();
    for (const token of parsed.tokens ?? []) {
      if (token.kind !== "option" || !token.name) continue;
      if (seen.has(token.name) && !["include", "exclude"].includes(token.name))
        throw new AgentDocxError(
          "INVALID_ARGUMENT",
          `Duplicate option: --${token.name}`,
        );
      seen.add(token.name);
    }
    return parsed;
  } catch (error) {
    if (error instanceof AgentDocxError) throw error;
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      error instanceof Error ? error.message : String(error),
    );
  }
};

export type CliOptionValues = Record<
  string,
  string | boolean | readonly string[] | undefined
>;

export type CliCommand =
  | { mode: "help" }
  | { mode: "version" }
  | { mode: "mcp" }
  | { mode: "profiles"; json: boolean }
  | { mode: "inspect"; path: string; json: boolean }
  | { mode: "batch-files"; paths: readonly string[]; values: CliOptionValues }
  | { mode: "batch-jsonl"; values: CliOptionValues }
  | { mode: "watch"; path: string; values: CliOptionValues }
  | {
      mode: "single";
      input:
        | { kind: "file"; path: string }
        | { kind: "stdin"; explicit: boolean };
      values: CliOptionValues;
    }
  | {
      mode: "workflow";
      command:
        | "project"
        | "document"
        | "revision"
        | "draft"
        | "review"
        | "validate"
        | "export"
        | "import"
        | "import-redline"
        | "filing-set"
        | "agent";
      args: readonly string[];
    };

const specs = {
  help: { type: "boolean" },
  version: { type: "boolean" },
  config: { type: "string" },
  profile: { type: "string" },
  "filing-kind": { type: "string" },
  template: { type: "string" },
  "page-size": { type: "string" },
  "page-width-in": { type: "string" },
  "page-height-in": { type: "string" },
  "margin-in": { type: "string" },
  "margin-top-in": { type: "string" },
  "margin-right-in": { type: "string" },
  "margin-bottom-in": { type: "string" },
  "margin-left-in": { type: "string" },
  "font-family": { type: "string" },
  "font-regular": { type: "string" },
  "font-bold": { type: "string" },
  "font-italic": { type: "string" },
  "font-bold-italic": { type: "string" },
  "font-size-pt": { type: "string" },
  "line-spacing": { type: "string" },
  "page-limit": { type: "string" },
  "fail-over-limit": { type: "boolean" },
  paragraphs: { type: "boolean" },
  sections: { type: "boolean" },
  trim: { type: "boolean" },
  "trim-limit": { type: "string" },
  "trim-threshold": { type: "string" },
  renderer: { type: "string" },
  "office-timeout": { type: "string" },
  "libreoffice-path": { type: "string" },
  json: { type: "boolean" },
  output: { type: "string" },
  batch: { type: "boolean" },
  "input-jsonl": { type: "boolean" },
  recursive: { type: "boolean" },
  "no-recursive": { type: "boolean" },
  include: { type: "string", multiple: true },
  exclude: { type: "string", multiple: true },
  watch: { type: "boolean" },
  jsonl: { type: "boolean" },
  "debounce-ms": { type: "string" },
  poll: { type: "boolean" },
} as const;

export const cliHelp = `Usage:
  agent-docx --help
  agent-docx --version
  agent-docx measure [FILE.md|-] [options]
  agent-docx profiles [--json]
  agent-docx template inspect FILE.docx [--json]

Project workflow:
  agent-docx project init|add ...
  agent-docx document configure ...
  agent-docx revision checkpoint|list|show|restore|diff|resolve ...
  agent-docx draft guidance|evaluate|apply ...
  agent-docx review add|resolve ...
  agent-docx validate ...
  agent-docx export ... [--mode clean|redline|pdf]
  agent-docx import ...
  agent-docx import-redline ...
  agent-docx filing-set add|remove|get|validate ...
  agent-docx agent --input-jsonl
  agent-docx agent --watch --project FILE --document ID --jsonl

MCP:
  agent-docx mcp
  Serves the version-1 agent protocol as a Model Context Protocol server
  over stdio (newline-delimited JSON-RPC).

Measure options retain batch, watch, layout, diagnostics, renderer, and --output support after the measure command.
Machine output is JSON/JSONL on stdout; fatal records are JSON on stderr.
`;

function parseMeasureArgs(args: readonly string[]): CliCommand {
  const parsed = parseCliArgsStrict({
    args: [...args],
    options: specs,
    strict: true,
    allowPositionals: true,
    tokens: true,
  });
  const values = parsed.values as CliOptionValues;


  if (values.help || values.version) {
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      "Use an explicit command; --help and --version are global forms only",
    );
  }
  if (
    values.output === "" ||
    values.output === "-" ||
    (typeof values.output === "string" && values.output.length === 0)
  ) {
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      "--output requires a file path other than -",
    );
  }
  if (values.recursive && values["no-recursive"]) {
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      "--recursive and --no-recursive cannot be combined",
    );
  }
  const hasDiscoveryOptions =
    values.recursive === true ||
    values["no-recursive"] === true ||
    values.include !== undefined ||
    values.exclude !== undefined;

  if (values.batch) {
    if (values.json || values.jsonl || values.watch || values.output) {
      throw new AgentDocxError(
        "INVALID_ARGUMENT",
        "Invalid batch option combination",
      );
    }
    if (values["input-jsonl"]) {
      if (hasDiscoveryOptions) {
        throw new AgentDocxError(
          "INVALID_ARGUMENT",
          "Discovery options require positional batch",
        );
      }
      if (parsed.positionals.length) {
        throw new AgentDocxError(
          "INVALID_ARGUMENT",
          "JSONL batch requires non-TTY stdin and no positionals",
        );
      }
      return { mode: "batch-jsonl", values };
    }
    if (!parsed.positionals.length) {
      throw new AgentDocxError(
        "INVALID_ARGUMENT",
        "Positional batch requires file paths",
      );
    }
    if (parsed.positionals.includes("-")) {
      throw new AgentDocxError(
        "INVALID_ARGUMENT",
        "Positional batch does not accept stdin; use --input-jsonl",
      );
    }
    if (parsed.positionals.some((token) => token.length === 0)) {
      throw new AgentDocxError(
        "INVALID_ARGUMENT",
        "Batch selector must not be empty",
      );
    }
    return { mode: "batch-files", paths: parsed.positionals, values };
  }

  if (values.watch) {
    if (
      parsed.positionals.length !== 1 ||
      values.json ||
      values["input-jsonl"] ||
      values.output ||
      hasDiscoveryOptions
    ) {
      throw new AgentDocxError(
        "INVALID_ARGUMENT",
        "Watch requires exactly one file and no batch/JSON/output/discovery options",
      );
    }
    return { mode: "watch", path: parsed.positionals[0]!, values };
  }

  if (hasDiscoveryOptions) {
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      "Discovery options require positional batch",
    );
  }
  if (values.output !== undefined && typeof values.output !== "string") {
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      "--output requires a file path",
    );
  }
  if (values.jsonl || values["debounce-ms"] !== undefined || values.poll) {
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      "Watch-only option used outside watch",
    );
  }
  if (parsed.positionals.length > 1) {
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      "Single mode accepts at most one input",
    );
  }
  const token = parsed.positionals[0];
  return {
    mode: "single",
    input:
      token === undefined || token === "-"
        ? { kind: "stdin", explicit: token === "-" }
        : { kind: "file", path: token },
    values,
  };
}

type WorkflowCommandName = Extract<CliCommand, { mode: "workflow" }>["command"];

const workflowSubcommands: Record<WorkflowCommandName, readonly string[]> = {
  project: ["init", "add"],
  document: ["configure"],
  revision: ["checkpoint", "list", "show", "restore", "diff", "resolve"],
  draft: ["guidance", "evaluate", "apply"],
  review: ["add", "resolve"],
  validate: [],
  export: [],
  import: [],
  "import-redline": [],
  "filing-set": ["add", "remove", "get", "validate"],
  agent: [],
};

export function parseCliArgs(args: readonly string[]): CliCommand {
  if (args.length === 1 && args[0] === "--help") return { mode: "help" };
  if (args.length === 1 && args[0] === "--version") return { mode: "version" };
  if (args.includes("--help") || args.includes("--version")) {
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      "--help and --version must be used alone",
    );
  }

  const [command, ...rest] = args;
  if (!command || command.startsWith("-")) {
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      "Expected an explicit command; use agent-docx --help",
    );
  }
  if (command === "measure") return parseMeasureArgs(rest);
  if (command === "profiles") {
    const parsed = parseCliArgsStrict({
      args: rest,
      options: specs,
      strict: true,
      allowPositionals: true,
    });
    const values = parsed.values as CliOptionValues;
    if (
      parsed.positionals.length ||
      Object.keys(values).some((key) => key !== "json")
    ) {
      throw new AgentDocxError(
        "INVALID_ARGUMENT",
        "profiles accepts only optional --json",
      );
    }
    return { mode: "profiles", json: values.json === true };
  }
  if (command === "template") {
    if (rest[0] !== "inspect") {
      throw new AgentDocxError(
        "INVALID_ARGUMENT",
        "template requires the inspect subcommand",
      );
    }
    const parsed = parseCliArgsStrict({
      args: rest.slice(1),
      options: specs,
      strict: true,
      allowPositionals: true,
    });
    const values = parsed.values as CliOptionValues;
    if (
      parsed.positionals.length !== 1 ||
      Object.keys(values).some((key) => key !== "json")
    ) {
      throw new AgentDocxError(
        "INVALID_ARGUMENT",
        "template inspect requires one DOCX path and optional --json",
      );
    }
    return {
      mode: "inspect",
      path: parsed.positionals[0]!,
      json: values.json === true,
    };
  }

  if (command === "mcp") {
    if (rest.length > 0)
      throw new AgentDocxError("INVALID_ARGUMENT", "mcp accepts no arguments");
    return { mode: "mcp" };
  }

  if (!(command in workflowSubcommands)) {
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `Unknown command: ${command}; use agent-docx --help`,
    );
  }
  const workflow = command as WorkflowCommandName;
  const allowed = workflowSubcommands[workflow];
  if (
    allowed.length > 0 &&
    (rest[0] === undefined || !allowed.includes(rest[0]))
  ) {
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `${workflow} requires one of: ${allowed.join(", ")}`,
    );
  }
  return { mode: "workflow", command: workflow, args: rest };
}
