import { parseArgs } from "node:util";
import { AgentDocxError } from "./types.js";

export type CliOptionValues = Record<
  string,
  string | boolean | readonly string[] | undefined
>;

export type CliCommand =
  | { mode: "help" }
  | { mode: "version" }
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
  "inspect-template": { type: "boolean" },
} as const;

export const cliHelp = `Usage: agent-docx [options] [FILE.md|-]\n\nEstimate DOCX-equivalent pages for legal Markdown.\n\nModes: --inspect-template FILE.docx, --batch FILE..., --batch --input-jsonl, --watch FILE\nOutput: --json (single/inspect), --jsonl (watch), --output FILE.docx (single)\nBatch discovery: --recursive, --no-recursive, --include GLOB, --exclude GLOB\n`;

export function parseCliArgs(args: readonly string[]): CliCommand {
  const parsed = parseArgs({
    args: [...args],
    options: specs,
    strict: true,
    allowPositionals: true,
    tokens: true,
  });
  const values = parsed.values as CliOptionValues;

  const seen = new Set<string>();
  for (const token of parsed.tokens) {
    if (token.kind !== "option" || !token.name) continue;
    if (seen.has(token.name) && !["include", "exclude"].includes(token.name)) {
      throw new AgentDocxError(
        "INVALID_ARGUMENT",
        `Duplicate option: --${token.name}`,
      );
    }
    seen.add(token.name);
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

  if (values.help || values.version) {
    const selected = values.help ? "help" : "version";
    if (
      parsed.positionals.length ||
      Object.keys(values).some((key) => key !== selected)
    ) {
      throw new AgentDocxError(
        "INVALID_ARGUMENT",
        "--help and --version must be used alone",
      );
    }
    return { mode: selected };
  }

  if (values["inspect-template"]) {
    if (
      parsed.positionals.length !== 1 ||
      Object.keys(values).some(
        (key) => !["inspect-template", "json"].includes(key),
      )
    ) {
      throw new AgentDocxError(
        "INVALID_ARGUMENT",
        "Inspect mode requires exactly one DOCX path and optional --json",
      );
    }
    return {
      mode: "inspect",
      path: parsed.positionals[0]!,
      json: values.json === true,
    };
  }

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
        "Watch requires exactly one file and no batch/JSON/inspect options",
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
