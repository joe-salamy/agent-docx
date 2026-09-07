import { link, open, unlink } from "node:fs/promises";
import { resolve } from "node:path";
import { cliHelp, parseCliArgs } from "./cli-args.js";
import { readInputFile } from "./input.js";
import { jsonlLines, strictUtf8 } from "./jsonl.js";
import { AgentDocxError } from "./types.js";
import type { MeasurementResult } from "./measurement.js";
import { serializableMeasurement } from "./measurement.js";
import type {
  CliJsonlRequest,
  CliRuntime,
  CliSequenceState,
  CliTrigger,
  OutputFileHandle,
  OutputFileIo,
  Source,
} from "./cli-contract.js";
import { resolveBatchInputs } from "./cli/run-batch.js";
import { asciiInteger, optionsFrom } from "./cli/run-options.js";
import {
  agentFatalRecord,
  endRecord,
  errorRecord,
  errorStatus,
  fatalRecord,
  human,
  humanProfiles,
  profileCatalog,
  readyRecord,
  resultRecord,
} from "./cli/run-report.js";
// runWatchController is dynamically imported per-command (see watch handling).

export type {
  BatchSelection,
  CliErrorPayload,
  CliErrorRecord,
  CliFatalRecord,
  CliJsonlRequest,
  CliResultRecord,
  CliRuntime,
  CliSequenceState,
  CliSource,
  CliTrigger,
  CliWatchEndRecord,
  CliWatchReadyRecord,
  OutputFileHandle,
  OutputFileIo,
  SerializableConfig,
  Source,
} from "./cli-contract.js";

type SequenceState = CliSequenceState;
let outputStageSequence = 0;
const outputFileIo: OutputFileIo = { open, link, unlink };

const isFsCode = (error: unknown, code: string): boolean =>
  error !== null &&
  typeof error === "object" &&
  "code" in error &&
  error.code === code;

const outputWriteError = (
  displayPath: string,
  error: unknown,
): AgentDocxError =>
  new AgentDocxError(
    "OUTPUT_WRITE_FAILED",
    `Failed to write output: ${displayPath}`,
    { cause: error instanceof Error ? error.message : String(error) },
  );

export async function writeOutputExclusive(
  resolvedPath: string,
  displayPath: string,
  bytes: Uint8Array,
  io: OutputFileIo = outputFileIo,
): Promise<void> {
  let stagePath: string | undefined;
  let handle: OutputFileHandle | undefined;
  try {
    for (let attempt = 0; attempt < 1000; attempt++) {
      const sequence = ++outputStageSequence;
      stagePath = `${resolvedPath}.agent-docx-stage-${process.pid}-${sequence}`;
      try {
        handle = await io.open(stagePath, "wx");
        break;
      } catch (error) {
        if (!isFsCode(error, "EEXIST")) throw error;
      }
    }
    if (!handle || !stagePath)
      throw new Error("Could not allocate a unique output stage file");
    await handle.writeFile(bytes);
    await handle.sync?.();
    await handle.close();
    handle = undefined;
    const publish = io.link ?? link;
    try {
      await publish(stagePath, resolvedPath);
    } catch (error) {
      if (isFsCode(error, "EEXIST"))
        throw new AgentDocxError(
          "OUTPUT_EXISTS",
          `Output already exists: ${displayPath}`,
        );
      throw error;
    }
    await io.unlink(stagePath);
    stagePath = undefined;
  } catch (error) {
    if (handle) {
      try {
        await handle.close();
      } catch {}
    }
    if (stagePath) {
      try {
        await io.unlink(stagePath);
      } catch {}
    }
    if (error instanceof AgentDocxError) throw error;
    throw outputWriteError(displayPath, error);
  }
}

export async function executeCli(
  args: readonly string[],
  runtime: CliRuntime,
  state: SequenceState,
): Promise<number> {
  const command = parseCliArgs(args);
  if (command.mode === "help") {
    await runtime.writeStdout(cliHelp);
    return 0;
  }
  if (command.mode === "version") {
    await runtime.writeStdout(`${runtime.version}\n`);
    return 0;
  }
  if (command.mode === "profiles") {
    const catalog = profileCatalog();
    await runtime.writeStdout(
      command.json ? `${JSON.stringify(catalog)}\n` : humanProfiles(catalog),
    );
    return 0;
  }
  if (command.mode === "inspect") {
    // Lazy docx inspect: avoids yauzl/saxes load for --help
    const { inspectDocxTemplate } = await import("./docx/inspect.js");
    const result = await inspectDocxTemplate(
      await readInputFile(resolve(runtime.cwd, command.path), "DOCX template"),
    );
    await runtime.writeStdout(
      command.json
        ? `${JSON.stringify(result)}\n`
        : `Template: ${result.package.sha256}\nSections: ${result.sections.length}; selected ${result.selectedSection}\n`,
    );
    return 0;
  }

  if (command.mode === "mcp") {
    // Lazy MCP: only for mcp server mode
    const { runMcpServer } = await import("./mcp.js");
    return runMcpServer(runtime);
  }

  if (command.mode === "skills") {
    // Lazy skills: avoid fs/cp overhead for --help
    const { listSkills, installSkills } = await import("./skills.js");
    if (command.subcommand === "list") {
      const skills = await listSkills();
      if (command.values.json === true) {
        await runtime.writeStdout(`${JSON.stringify(skills, null, 2)}\n`);
      } else {
        if (skills.length === 0) {
          await runtime.writeStdout("No skills available\n");
        } else {
          for (const s of skills) {
            await runtime.writeStdout(
              `${s.name}@${s.version}: ${s.description}\n`,
            );
          }
        }
      }
      return 0;
    }
    const dest =
      typeof command.values.dest === "string" ? command.values.dest : undefined;
    const global = command.values.global === true;
    const force = command.values.force === true;
    const dryRun = command.values["dry-run"] === true;
    const json = command.values.json === true;
    const installOpts: {
      cwd: string;
      dest?: string;
      global?: boolean;
      force?: boolean;
      dryRun?: boolean;
    } = {
      cwd: runtime.cwd,
    };
    if (dest !== undefined) installOpts.dest = dest;
    if (global) installOpts.global = true;
    if (force) installOpts.force = true;
    if (dryRun) installOpts.dryRun = true;
    const result = await installSkills(installOpts);
    const skipped = result.results.filter((r) => r.status === "skipped");
    if (skipped.length > 0 && !json) {
      for (const r of skipped) {
        await runtime.writeStderr(
          `Skipped ${r.name}: exists at ${r.destPath} (use --force to overwrite)\n`,
        );
      }
    }
    if (json) {
      await runtime.writeStdout(
        `${JSON.stringify(
          {
            destBase: result.destBase,
            dryRun,
            results: result.results,
          },
          null,
          2,
        )}\n`,
      );
    } else {
      for (const r of result.results) {
        if (r.status === "skipped") continue;
        const verb =
          r.status === "dry-run"
            ? "would install"
            : r.status === "overwritten"
              ? "overwrote"
              : "installed";
        await runtime.writeStdout(`${verb} ${r.name} -> ${r.destPath}\n`);
      }
      if (dryRun) {
        await runtime.writeStdout(
          `dry-run: no files written (destBase: ${result.destBase})\n`,
        );
      }
    }
    if (skipped.length > 0) {
      throw new AgentDocxError(
        "OUTPUT_EXISTS",
        `Skills already exist at ${skipped[0]!.destPath} (use --force)`,
      );
    }
    return 0;
  }

  if (command.mode === "workflow") {
    // Lazy workflow: project/store/proper-lockfile heavy
    const { runWorkflowCommand } = await import("./cli-workflow.js");
    return runWorkflowCommand(command, runtime, state);
  }

  if (command.mode === "batch-files" || command.mode === "batch-jsonl") {
    // Lazy renderers: unified/remark/fontkit heavy, only for measure modes
    const { measureMarkdown } = await import("./renderers/index.js");
    const values = command.values;
    const base = await optionsFrom(values, runtime.cwd);
    const batchSources =
      command.mode === "batch-files"
        ? await resolveBatchInputs(
            command.paths,
            {
              recursive:
                values.recursive === true
                  ? true
                  : values["no-recursive"] === true
                    ? false
                    : (base.batch?.recursive ?? true),
              include: Array.isArray(values.include)
                ? values.include
                : (base.batch?.include ?? ["*.md"]),
              exclude: Array.isArray(values.exclude)
                ? values.exclude
                : (base.batch?.exclude ?? []),
            },
            runtime.cwd,
          )
        : [];
    let failed = false;
    let over = false;
    if (command.mode === "batch-jsonl") {
      if (runtime.stdinIsTTY) {
        throw new AgentDocxError(
          "INVALID_ARGUMENT",
          "JSONL batch requires non-TTY stdin and no positionals",
        );
      }
      for await (const line of jsonlLines(runtime)) {
        if (!line.trim()) continue;
        let requestId: string | number | null = null;
        let source: Source = { kind: "stdin" };
        try {
          let request: unknown;
          try {
            request = JSON.parse(line);
          } catch (error) {
            throw new AgentDocxError(
              "INVALID_ARGUMENT",
              error instanceof Error ? error.message : "Invalid JSON",
            );
          }
          if (
            !request ||
            typeof request !== "object" ||
            Array.isArray(request)
          ) {
            throw new AgentDocxError(
              "INVALID_ARGUMENT",
              "request must be an object",
            );
          }
          const record = request as CliJsonlRequest;
          const hasPath = "path" in record && typeof record.path === "string";
          const hasMarkdown =
            "markdown" in record && typeof record.markdown === "string";
          if (hasPath && "path" in record && record.path !== "") {
            source = {
              kind: "file",
              path: record.path,
              resolvedPath: resolve(runtime.cwd, record.path),
            };
          } else if (hasMarkdown) {
            source = {
              kind: "inline",
              name:
                "name" in record && typeof record.name === "string"
                  ? record.name
                  : null,
            };
          }
          if ("id" in record) {
            if (
              record.id !== null &&
              typeof record.id !== "string" &&
              !(typeof record.id === "number" && Number.isFinite(record.id))
            ) {
              throw new AgentDocxError(
                "INVALID_ARGUMENT",
                "id must be a string, finite number, or null",
              );
            }
            requestId = record.id;
          }
          if (hasPath === hasMarkdown) {
            throw new AgentDocxError(
              "INVALID_ARGUMENT",
              "exactly one of path or markdown is required",
            );
          }
          if (hasPath && "path" in record && record.path === "") {
            throw new AgentDocxError(
              "INVALID_ARGUMENT",
              "path must not be empty",
            );
          }
          if (
            hasMarkdown &&
            "name" in record &&
            record.name !== undefined &&
            typeof record.name !== "string"
          ) {
            throw new AgentDocxError(
              "INVALID_ARGUMENT",
              "name must be a string",
            );
          }
          const allowed: readonly string[] = hasPath
            ? ["id", "path"]
            : ["id", "name", "markdown"];
          const unknown = Object.keys(record).find(
            (key) => !allowed.includes(key),
          );
          if (unknown) {
            throw new AgentDocxError(
              "INVALID_ARGUMENT",
              `unknown request key: ${unknown}`,
            );
          }
          const markdown =
            hasPath && "path" in record
              ? await strictUtf8(
                  await readInputFile(
                    resolve(runtime.cwd, record.path),
                    "Markdown input",
                  ),
                )
              : "markdown" in record
                ? record.markdown
                : "";
          const rawMeasurementJsonl = await measureMarkdown(
            markdown,
            base.options,
          );
          const measurement =
            base.linePageFilter && rawMeasurementJsonl.deterministic.lines
              ? {
                  ...rawMeasurementJsonl,
                  deterministic: {
                    ...rawMeasurementJsonl.deterministic,
                    lines: rawMeasurementJsonl.deterministic.lines.filter((l) =>
                      base.linePageFilter!.includes(l.page),
                    ),
                  },
                }
              : rawMeasurementJsonl;
          await runtime.writeStdout(
            `${JSON.stringify(
              resultRecord(
                state,
                "batch",
                source,
                measurement,
                runtime.cwd,
                requestId,
              ),
            )}\n`,
          );
          if (
            values["fail-over-limit"] &&
            measurement.budget &&
            !measurement.budget.withinLimit
          ) {
            over = true;
          }
        } catch (error) {
          failed = true;
          await runtime.writeStdout(
            `${JSON.stringify(
              errorRecord(
                state,
                "batch",
                source,
                error,
                runtime.cwd,
                requestId,
              ),
            )}\n`,
          );
        }
      }
    } else {
      for (const source of batchSources) {
        const resolvedPath = source.resolvedPath;
        try {
          const rawMeasurementBatch = await measureMarkdown(
            await strictUtf8(
              await readInputFile(resolvedPath, "Markdown input"),
            ),
            base.options,
          );
          const measurement =
            base.linePageFilter && rawMeasurementBatch.deterministic.lines
              ? {
                  ...rawMeasurementBatch,
                  deterministic: {
                    ...rawMeasurementBatch.deterministic,
                    lines: rawMeasurementBatch.deterministic.lines.filter((l) =>
                      base.linePageFilter!.includes(l.page),
                    ),
                  },
                }
              : rawMeasurementBatch;
          await runtime.writeStdout(
            `${JSON.stringify(
              resultRecord(state, "batch", source, measurement, runtime.cwd),
            )}\n`,
          );
          if (
            values["fail-over-limit"] &&
            measurement.budget &&
            !measurement.budget.withinLimit
          ) {
            over = true;
          }
        } catch (error) {
          failed = true;
          await runtime.writeStdout(
            `${JSON.stringify(
              errorRecord(state, "batch", source, error, runtime.cwd),
            )}\n`,
          );
        }
      }
    }
    return failed ? 1 : over ? 3 : 0;
  }

  if (command.mode === "watch") {
    // Lazy watch/renderers: chokidar + fontkit heavy
    const { measureMarkdown } = await import("./renderers/index.js");
    const { runWatchController } = await import("./watch.js");
    const values = command.values;
    const token = command.path;
    const path = resolve(runtime.cwd, token);
    const source: Source = { kind: "file", path: token, resolvedPath: path };
    const debounce =
      typeof values["debounce-ms"] === "string"
        ? asciiInteger(values["debounce-ms"], "--debounce-ms", 0, 60000)
        : 75;
    const loaded = await optionsFrom(values, runtime.cwd);
    const dependencies = [path, ...loaded.dependencies];
    const toCliTrigger = (trigger: {
      kind: "initial" | "change";
      paths: readonly string[];
    }): CliTrigger => {
      if (trigger.kind === "initial")
        return { kind: "initial", paths: trigger.paths };
      const changed = trigger.paths[0] ?? path;
      return {
        kind: changed === path ? "source-change" : "dependency-change",
        paths: [resolve(changed)],
      };
    };
    const watchCode = await runWatchController<MeasurementResult>({
      watchPaths: dependencies,
      watchOptions: {
        atomic: 200,
        awaitWriteFinish: { stabilityThreshold: 200, pollInterval: 50 },
        usePolling: values.poll === true,
      },
      debounceMs: debounce,
      run: async (trigger) => {
        if (trigger.kind === "initial" && values.jsonl)
          await runtime.writeStdout(
            `${JSON.stringify(
              readyRecord(state, source, dependencies, runtime.cwd),
            )}\n`,
          );
        return measureMarkdown(
          await strictUtf8(await readInputFile(path, "Markdown input")),
          loaded.options,
        );
      },
      emitResult: async (measurement, trigger) => {
        const cliTrigger = toCliTrigger(trigger);
        const filteredMeasurement =
          loaded.linePageFilter && measurement.deterministic.lines
            ? {
                ...measurement,
                deterministic: {
                  ...measurement.deterministic,
                  lines: measurement.deterministic.lines.filter((l) =>
                    loaded.linePageFilter!.includes(l.page),
                  ),
                },
              }
            : measurement;
        await runtime.writeStdout(
          values.jsonl
            ? `${JSON.stringify(
                resultRecord(
                  state,
                  "watch",
                  source,
                  measurement,
                  runtime.cwd,
                  null,
                  cliTrigger,
                ),
              )}\n`
            : `\n[${cliTrigger.kind}]\n${human(filteredMeasurement, {
                paragraphs: values.paragraphs === true,
                trim: loaded.options.trim !== undefined,
                lines: values.lines === true,
                ...(loaded.linePageFilter
                  ? { linesPageFilter: loaded.linePageFilter }
                  : {}),
              })}`,
        );
      },
      emitError: async (error, trigger) => {
        const text = `${JSON.stringify(
          errorRecord(
            state,
            "watch",
            source,
            error,
            runtime.cwd,
            null,
            toCliTrigger(trigger),
          ),
        )}\n`;
        await (values.jsonl
          ? runtime.writeStdout(text)
          : runtime.writeStderr(text));
      },
      signal: (name, listener) => runtime.onceSignal(name, listener),
      onStop: async (reason) => {
        if (values.jsonl)
          await runtime.writeStdout(
            `${JSON.stringify(endRecord(state, source, reason, runtime.cwd))}\n`,
          );
      },
    });
    return watchCode;
  }

  const loaded = await optionsFrom(command.values, runtime.cwd);
  // Lazy renderers: only for single-file measure
  const { measureMarkdown } = await import("./renderers/index.js");
  const outputPath =
    typeof command.values.output === "string"
      ? command.values.output
      : undefined;
  if (outputPath) loaded.options.includeGeneratedDocx = true;
  let markdown: string;
  if (command.input.kind === "file") {
    markdown = await strictUtf8(
      await readInputFile(
        resolve(runtime.cwd, command.input.path),
        "Markdown input",
      ),
    );
  } else {
    if (!command.input.explicit && runtime.stdinIsTTY) {
      throw new AgentDocxError(
        "INVALID_ARGUMENT",
        "No input: provide a file or pipe Markdown on stdin",
      );
    }
    markdown = await strictUtf8(await runtime.readStdin());
  }
  const rawMeasurementSingle = await measureMarkdown(markdown, loaded.options);
  const measurement =
    loaded.linePageFilter && rawMeasurementSingle.deterministic.lines
      ? {
          ...rawMeasurementSingle,
          deterministic: {
            ...rawMeasurementSingle.deterministic,
            lines: rawMeasurementSingle.deterministic.lines.filter((l) =>
              loaded.linePageFilter!.includes(l.page),
            ),
          },
        }
      : rawMeasurementSingle;
  if (outputPath) {
    await writeOutputExclusive(
      resolve(runtime.cwd, outputPath),
      outputPath,
      measurement.generatedDocx!,
    );
  }
  if (command.values.json) {
    await runtime.writeStdout(
      `${JSON.stringify(serializableMeasurement(measurement))}\n`,
    );
  } else {
    await runtime.writeStdout(
      human(measurement, {
        paragraphs: command.values.paragraphs === true,
        trim: loaded.options.trim !== undefined,
        lines: command.values.lines === true,
        ...(loaded.linePageFilter
          ? { linesPageFilter: loaded.linePageFilter }
          : {}),
      }),
    );
    for (const warning of measurement.deterministic.warnings) {
      await runtime.writeStderr(`${warning.code}: ${warning.message}\n`);
    }
  }
  if (command.values["fail-over-limit"]) {
    const budgets =
      measurement.mode === "compare"
        ? Object.values(measurement.budgetBySource ?? {})
        : [measurement.budget];
    if (budgets.some((budget) => budget && !budget.withinLimit)) return 3;
  }
  return 0;
}

export async function runCli(
  args: readonly string[],
  runtime: CliRuntime,
): Promise<number> {
  const state = { sequence: 0 };
  try {
    return await executeCli(args, runtime, state);
  } catch (error) {
    await runtime.writeStderr(
      `${JSON.stringify(
        args[0] === "agent"
          ? agentFatalRecord(error, state.sequence + 1)
          : fatalRecord(error),
      )}\n`,
    );
    return errorStatus(error);
  }
}
