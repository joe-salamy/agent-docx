import {
  glob,
  lstat,
  open,
  readFile,
  realpath,
  stat,
  unlink,
} from "node:fs/promises";
import type { Stats } from "node:fs";
import { dirname, isAbsolute, relative, resolve, sep } from "node:path";
import { fileURLToPath } from "node:url";
import { Ajv2020 } from "ajv/dist/2020.js";
import chokidar from "chokidar";
import { parseCliArgs, cliHelp, type CliOptionValues } from "./cli-args.js";
import { inspectDocxTemplate } from "./docx/inspect.js";
import { measureMarkdown } from "./renderers/index.js";
import {
  MdPageCountError,
  type EstimateOptions,
  type FontSetInput,
  type LayoutOverrides,
  type MeasureOptions,
  type MeasurementResult,
  type RendererMode,
} from "./types.js";

export interface CliRuntime {
  readonly cwd: string;
  readonly stdinIsTTY: boolean;
  readonly version: string;
  readStdin(): Promise<Uint8Array>;
  writeStdout(text: string): Promise<void>;
  writeStderr(text: string): Promise<void>;
  onceSignal(signal: "SIGINT" | "SIGTERM", listener: () => void): void;
}

type SerializableConfig = {
  profile?: "us-district-conventional" | "frap-32" | "cand-civil";
  templatePath?: string;
  layout?: LayoutOverrides;
  fontSet?: {
    family: string;
    regularPath: string;
    boldPath?: string;
    italicPath?: string;
    boldItalicPath?: string;
  };
  filingKind?: EstimateOptions["filingKind"];
  pageLimit?: number;
  paragraphDiagnostics?: boolean;
  sectionDiagnostics?: boolean;
  trim?: EstimateOptions["trim"];
  renderer?: RendererMode;
  officeTimeoutMs?: number;
  word?: { powerShellPath?: string };
  libreoffice?: {
    executablePath?: string;
    installedFonts?: { family: string; path: string }[];
  };
  batch?: {
    recursive?: boolean;
    include?: string[];
    exclude?: string[];
  };
};

type Source =
  | { kind: "file"; path: string; resolvedPath: string }
  | { kind: "stdin" }
  | { kind: "inline"; name: string | null };

type BatchSelection = {
  readonly recursive: boolean;
  readonly include: readonly string[];
  readonly exclude: readonly string[];
};

function invalidPattern(pattern: string) {
  if (!pattern || isAbsolute(pattern) || pattern.startsWith("!")) return true;
  let brackets = 0;
  let escaped = false;
  for (const character of pattern) {
    if (escaped) {
      escaped = false;
    } else if (character === "\\") {
      escaped = true;
    } else if (character === "[") {
      brackets++;
    } else if (character === "]") {
      if (brackets === 0) return true;
      brackets--;
    }
  }
  return brackets !== 0;
}

const normalizedRelativePath = (cwd: string, path: string) =>
  relative(cwd, path).split(sep).join("/");

export async function resolveBatchInputs(
  selectors: readonly string[],
  selection: BatchSelection,
  cwd: string,
): Promise<Extract<Source, { kind: "file" }>[]> {
  for (const pattern of [...selection.include, ...selection.exclude]) {
    if (invalidPattern(pattern)) {
      throw new MdPageCountError(
        "INVALID_ARGUMENT",
        `Invalid batch pattern: ${pattern}`,
      );
    }
  }
  const output: Extract<Source, { kind: "file" }>[] = [];
  const seen = new Set<string>();
  const add = async (
    source: Extract<Source, { kind: "file" }>,
    existing: boolean,
  ) => {
    const identity = existing
      ? await realpath(source.resolvedPath)
      : source.resolvedPath;
    if (seen.has(identity)) return;
    seen.add(identity);
    output.push(source);
  };
  const discoveredSources = async (
    pattern: string | readonly string[],
    root: string,
    excludes: readonly string[],
    errorPattern: string,
    directOnly: boolean,
  ) => {
    const candidates: Extract<Source, { kind: "file" }>[] = [];
    try {
      for await (const matched of glob(pattern, {
        cwd: root,
        exclude: excludes,
      })) {
        const resolvedPath = resolve(root, matched);
        const sourcePath = normalizedRelativePath(cwd, resolvedPath);
        const matchedPath = String(matched).split(sep).join("/");
        if (directOnly && matchedPath.includes("/")) continue;
        try {
          if (!(await stat(resolvedPath)).isFile()) continue;
        } catch {
          continue;
        }
        candidates.push({
          kind: "file",
          path: sourcePath,
          resolvedPath,
        });
      }
    } catch {
      throw new MdPageCountError(
        "INVALID_ARGUMENT",
        `Invalid batch pattern: ${errorPattern}`,
      );
    }
    candidates.sort((a, b) => (a.path < b.path ? -1 : a.path > b.path ? 1 : 0));
    const unique: Extract<Source, { kind: "file" }>[] = [];
    const localSeen = new Set<string>();
    for (const candidate of candidates) {
      const identity = await realpath(candidate.resolvedPath);
      if (localSeen.has(identity)) continue;
      localSeen.add(identity);
      unique.push(candidate);
    }
    return unique;
  };

  for (const selector of selectors) {
    if (!selector) {
      throw new MdPageCountError(
        "INVALID_ARGUMENT",
        "Batch selector must not be empty",
      );
    }
    if (selector === "-") {
      throw new MdPageCountError(
        "INVALID_ARGUMENT",
        "Positional batch does not accept stdin; use --input-jsonl",
      );
    }
    const resolvedSelector = resolve(cwd, selector);
    let information: Stats | undefined;
    try {
      information = await lstat(resolvedSelector);
    } catch (error) {
      if (
        !(
          error &&
          typeof error === "object" &&
          "code" in error &&
          error.code === "ENOENT"
        )
      ) {
        throw error;
      }
    }
    if (information) {
      if (information.isSymbolicLink()) {
        try {
          if ((await stat(resolvedSelector)).isFile()) {
            await add(
              {
                kind: "file",
                path: selector,
                resolvedPath: resolvedSelector,
              },
              true,
            );
            continue;
          }
        } catch {}
        throw new MdPageCountError(
          "INVALID_ARGUMENT",
          `Unsupported batch input: ${selector}`,
        );
      }
      if (information.isFile()) {
        await add(
          { kind: "file", path: selector, resolvedPath: resolvedSelector },
          true,
        );
        continue;
      }
      if (information.isDirectory()) {
        const include = selection.recursive
          ? selection.include.map((pattern) =>
              pattern.includes("/") ? pattern : `**/${pattern}`,
            )
          : selection.include.filter((pattern) => !pattern.includes("/"));
        const exclude = selection.exclude.map((pattern) =>
          selection.recursive && !pattern.includes("/")
            ? `**/${pattern}`
            : pattern,
        );
        const matches =
          include.length === 0
            ? []
            : await discoveredSources(
                include,
                resolvedSelector,
                exclude,
                include[0]!,
                !selection.recursive,
              );
        if (matches.length === 0) {
          throw new MdPageCountError(
            "INVALID_ARGUMENT",
            `Batch selector matched no files: ${selector}`,
          );
        }
        for (const source of matches) await add(source, true);
        continue;
      }
      throw new MdPageCountError(
        "INVALID_ARGUMENT",
        `Unsupported batch input: ${selector}`,
      );
    }

    if (selector.startsWith("!")) {
      throw new MdPageCountError(
        "INVALID_ARGUMENT",
        `Invalid batch pattern: ${selector}`,
      );
    }
    if (/[*?[\]]/.test(selector)) {
      if (invalidPattern(selector)) {
        throw new MdPageCountError(
          "INVALID_ARGUMENT",
          `Invalid batch pattern: ${selector}`,
        );
      }
      const exclude = selection.exclude.map((pattern) =>
        pattern.includes("/") ? pattern : `**/${pattern}`,
      );
      const matches = await discoveredSources(
        selector,
        cwd,
        exclude,
        selector,
        false,
      );
      if (matches.length === 0) {
        throw new MdPageCountError(
          "INVALID_ARGUMENT",
          `Batch selector matched no files: ${selector}`,
        );
      }
      for (const source of matches) await add(source, true);
      continue;
    }

    await add(
      { kind: "file", path: selector, resolvedPath: resolvedSelector },
      false,
    );
  }
  return output;
}
type SequenceState = { sequence: number };

export type OutputFileHandle = {
  writeFile(bytes: Uint8Array): Promise<void>;
  close(): Promise<void>;
};
export type OutputFileIo = {
  open(path: string, flags: "wx"): Promise<OutputFileHandle>;
  unlink(path: string): Promise<void>;
};

const outputFileIo: OutputFileIo = { open, unlink };

export async function writeOutputExclusive(
  resolvedPath: string,
  displayPath: string,
  bytes: Uint8Array,
  io: OutputFileIo = outputFileIo,
): Promise<void> {
  let handle: OutputFileHandle;
  try {
    handle = await io.open(resolvedPath, "wx");
  } catch (error) {
    if (
      error !== null &&
      typeof error === "object" &&
      "code" in error &&
      error.code === "EEXIST"
    ) {
      throw new MdPageCountError(
        "OUTPUT_EXISTS",
        `Output already exists: ${displayPath}`,
      );
    }
    throw new MdPageCountError(
      "OUTPUT_WRITE_FAILED",
      `Failed to write output: ${displayPath}`,
      { cause: error instanceof Error ? error.message : String(error) },
    );
  }
  try {
    await handle.writeFile(bytes);
    await handle.close();
  } catch (error) {
    try {
      await handle.close();
    } catch {}
    try {
      await io.unlink(resolvedPath);
    } catch {}
    throw new MdPageCountError(
      "OUTPUT_WRITE_FAILED",
      `Failed to write output: ${displayPath}`,
      { cause: error instanceof Error ? error.message : String(error) },
    );
  }
}

function asciiInteger(
  value: string,
  name: string,
  min = 1,
  max = Number.MAX_SAFE_INTEGER,
) {
  if (!/^(?:0|[1-9][0-9]*)$/.test(value)) {
    throw new MdPageCountError(
      "INVALID_ARGUMENT",
      `${name} requires ASCII digits`,
    );
  }
  const number = Number(value);
  if (!Number.isSafeInteger(number) || number < min || number > max) {
    throw new MdPageCountError("INVALID_ARGUMENT", `${name} is out of range`);
  }
  return number;
}

function decimal(value: string, name: string, positive = true) {
  if (!/^(?:0|[1-9][0-9]*)(?:[.][0-9]+)?$/.test(value)) {
    throw new MdPageCountError(
      "INVALID_ARGUMENT",
      `${name} has invalid decimal syntax`,
    );
  }
  const number = Number(value);
  if (!Number.isFinite(number) || (positive ? number <= 0 : number < 0)) {
    throw new MdPageCountError(
      "INVALID_ARGUMENT",
      `${name} must be ${positive ? "positive" : "nonnegative"}`,
    );
  }
  return number;
}

const twips = (number: number, scale: number) =>
  Math.floor(number * scale + 0.5);

async function strictUtf8(bytes: Uint8Array) {
  try {
    return new TextDecoder("utf-8", { fatal: true }).decode(bytes);
  } catch {
    throw new MdPageCountError("INPUT_NOT_UTF8", "Input is not valid UTF-8");
  }
}

async function loadConfig(pathToken: string): Promise<{
  config: SerializableConfig;
  base: string;
  path: string;
}> {
  const path = resolve(pathToken);
  let bytes: Uint8Array;
  try {
    bytes = await readFile(path);
  } catch {
    throw new MdPageCountError(
      "INPUT_NOT_FOUND",
      `Configuration not found: ${pathToken}`,
    );
  }
  let value: unknown;
  try {
    value = JSON.parse(await strictUtf8(bytes));
  } catch (error) {
    if (error instanceof MdPageCountError) throw error;
    throw new MdPageCountError(
      "INVALID_CONFIG",
      "Configuration is not valid JSON",
    );
  }
  const schemaPath = fileURLToPath(
    new URL("../config.schema.json", import.meta.url),
  );
  const schema = JSON.parse(await readFile(schemaPath, "utf8")) as object;
  const ajv = new Ajv2020({ allErrors: true, strict: true });
  const validate = ajv.compile(schema);
  if (!validate(value)) {
    throw new MdPageCountError(
      "INVALID_CONFIG",
      "Configuration does not match config.schema.json",
      { errors: JSON.stringify(validate.errors) },
    );
  }
  return {
    config: value as SerializableConfig,
    base: dirname(path),
    path,
  };
}

async function fileBytes(path: string, base: string) {
  return readFile(resolve(base, path));
}

async function optionsFrom(
  values: CliOptionValues,
  cwd: string,
): Promise<{
  options: MeasureOptions;
  dependencies: string[];
  batch: SerializableConfig["batch"];
}> {
  let config: SerializableConfig = {};
  let base = cwd;
  const dependencies: string[] = [];
  if (typeof values.config === "string") {
    const loaded = await loadConfig(resolve(cwd, values.config));
    config = loaded.config;
    base = loaded.base;
    dependencies.push(loaded.path);
  }
  const options: MeasureOptions = {};
  if (config.profile) options.profile = config.profile;
  if (config.filingKind) options.filingKind = config.filingKind;
  if (config.pageLimit) options.pageLimit = config.pageLimit;
  if (config.paragraphDiagnostics !== undefined) {
    options.paragraphDiagnostics = config.paragraphDiagnostics;
  }
  if (config.sectionDiagnostics !== undefined) {
    options.sectionDiagnostics = config.sectionDiagnostics;
  }
  if (config.trim !== undefined) options.trim = config.trim;
  if (config.renderer) options.renderer = config.renderer;
  if (config.officeTimeoutMs) options.officeTimeoutMs = config.officeTimeoutMs;
  if (config.word) {
    options.word = {
      ...config.word,
      ...(config.word.powerShellPath
        ? { powerShellPath: resolve(base, config.word.powerShellPath) }
        : {}),
    };
  }
  if (config.libreoffice) {
    options.libreoffice = {
      ...config.libreoffice,
      ...(config.libreoffice.executablePath
        ? { executablePath: resolve(base, config.libreoffice.executablePath) }
        : {}),
      ...(config.libreoffice.installedFonts
        ? {
            installedFonts: config.libreoffice.installedFonts.map((font) => ({
              family: font.family,
              path: resolve(base, font.path),
            })),
          }
        : {}),
    };
  }
  if (config.layout) options.layout = structuredClone(config.layout);

  const templateToken =
    typeof values.template === "string"
      ? resolve(cwd, values.template)
      : config.templatePath
        ? resolve(base, config.templatePath)
        : undefined;
  if (templateToken) {
    options.template = await inspectDocxTemplate(
      await readFile(templateToken),
      {
        fallbackProfile: options.profile,
      },
    );
    dependencies.push(templateToken);
  }

  const fontSpecification =
    typeof values["font-regular"] === "string"
      ? {
          family: String(values["font-family"] ?? ""),
          regularPath: resolve(cwd, values["font-regular"]),
          ...(typeof values["font-bold"] === "string"
            ? { boldPath: resolve(cwd, values["font-bold"]) }
            : {}),
          ...(typeof values["font-italic"] === "string"
            ? { italicPath: resolve(cwd, values["font-italic"]) }
            : {}),
          ...(typeof values["font-bold-italic"] === "string"
            ? { boldItalicPath: resolve(cwd, values["font-bold-italic"]) }
            : {}),
        }
      : config.fontSet
        ? {
            ...config.fontSet,
            regularPath: resolve(base, config.fontSet.regularPath),
            ...(config.fontSet.boldPath
              ? { boldPath: resolve(base, config.fontSet.boldPath) }
              : {}),
            ...(config.fontSet.italicPath
              ? { italicPath: resolve(base, config.fontSet.italicPath) }
              : {}),
            ...(config.fontSet.boldItalicPath
              ? { boldItalicPath: resolve(base, config.fontSet.boldItalicPath) }
              : {}),
          }
        : undefined;
  if (fontSpecification) {
    if (!fontSpecification.family) {
      throw new MdPageCountError(
        "INVALID_ARGUMENT",
        "--font-family is required with --font-* paths",
      );
    }
    const font: FontSetInput = {
      family: fontSpecification.family,
      regular: await fileBytes(fontSpecification.regularPath, "/"),
    };
    dependencies.push(fontSpecification.regularPath);
    if (fontSpecification.boldPath) {
      font.bold = await fileBytes(fontSpecification.boldPath, "/");
      dependencies.push(fontSpecification.boldPath);
    }
    if (fontSpecification.italicPath) {
      font.italic = await fileBytes(fontSpecification.italicPath, "/");
      dependencies.push(fontSpecification.italicPath);
    }
    if (fontSpecification.boldItalicPath) {
      font.boldItalic = await fileBytes(fontSpecification.boldItalicPath, "/");
      dependencies.push(fontSpecification.boldItalicPath);
    }
    options.fontSet = font;
  }

  if (typeof values.profile === "string") {
    options.profile = values.profile as MeasureOptions["profile"];
  }
  if (typeof values["filing-kind"] === "string") {
    options.filingKind = values["filing-kind"] as EstimateOptions["filingKind"];
  }
  if (typeof values["page-limit"] === "string") {
    options.pageLimit = asciiInteger(values["page-limit"], "--page-limit");
  }
  if (values.paragraphs === true) options.paragraphDiagnostics = true;
  if (values.sections === true) options.sectionDiagnostics = true;
  if (values.trim === true) options.trim = {};
  if (typeof values["trim-limit"] === "string") {
    options.trim = {
      ...(options.trim || {}),
      maxCandidates: asciiInteger(values["trim-limit"], "--trim-limit", 1, 100),
    };
  }
  if (typeof values["trim-threshold"] === "string") {
    options.trim = {
      ...(options.trim || {}),
      maxLastLineRatio: decimal(
        values["trim-threshold"],
        "--trim-threshold",
        false,
      ),
    };
  }
  if (typeof values.renderer === "string") {
    options.renderer = values.renderer as RendererMode;
  }
  if (typeof values["office-timeout"] === "string") {
    options.officeTimeoutMs = asciiInteger(
      values["office-timeout"],
      "--office-timeout",
      1000,
      600000,
    );
  }
  if (typeof values["libreoffice-path"] === "string") {
    options.libreoffice = {
      ...(options.libreoffice ?? {}),
      executablePath: resolve(cwd, values["libreoffice-path"]),
    };
  }

  const layout: LayoutOverrides = structuredClone(options.layout ?? {});
  const page = (layout.page ??= {});
  if (values["page-size"] === "letter") {
    Object.assign(page, { widthTwips: 12240, heightTwips: 15840 });
  } else if (values["page-size"] === "a4") {
    Object.assign(page, { widthTwips: 11907, heightTwips: 16839 });
  } else if (values["page-size"] !== undefined) {
    throw new MdPageCountError(
      "INVALID_ARGUMENT",
      "--page-size must be letter or a4",
    );
  }
  if (typeof values["page-width-in"] === "string") {
    page.widthTwips = twips(
      decimal(values["page-width-in"], "--page-width-in"),
      1440,
    );
  }
  if (typeof values["page-height-in"] === "string") {
    page.heightTwips = twips(
      decimal(values["page-height-in"], "--page-height-in"),
      1440,
    );
  }
  const edges = ["top", "right", "bottom", "left"] as const;
  if (typeof values["margin-in"] === "string") {
    const number = twips(
      decimal(values["margin-in"], "--margin-in", false),
      1440,
    );
    page.marginsTwips = {
      top: number,
      right: number,
      bottom: number,
      left: number,
    };
  }
  for (const edge of edges) {
    const raw = values[`margin-${edge}-in`];
    if (typeof raw === "string") {
      page.marginsTwips = {
        ...page.marginsTwips,
        [edge]: twips(decimal(raw, `--margin-${edge}-in`, false), 1440),
      };
    }
  }
  if (typeof values["font-size-pt"] === "string") {
    layout.body = {
      ...layout.body,
      fontSizeTwips: twips(
        decimal(values["font-size-pt"], "--font-size-pt"),
        20,
      ),
    };
  }
  if (typeof values["line-spacing"] === "string") {
    layout.body = {
      ...layout.body,
      lineSpacing: {
        rule: "auto",
        numerator: twips(
          decimal(values["line-spacing"], "--line-spacing"),
          240,
        ),
        denominator: 240,
      },
    };
  }
  if (Object.keys(layout).length) options.layout = layout;
  return { options, dependencies, batch: config.batch };
}

function serializableMeasurement(
  measurement: MeasurementResult,
): Omit<MeasurementResult, "generatedDocx"> {
  const { generatedDocx: _generatedDocx, ...serializable } = measurement;
  return serializable;
}

function human(measurement: MeasurementResult) {
  const deterministic = measurement.deterministic;
  const rows = [
    `Estimated pages: ${deterministic.pageCount} physical; ${deterministic.equivalentPages.toFixed(3)} equivalent`,
  ];
  if (deterministic.lastPage) {
    rows.push(
      `Last page: ${deterministic.lastPage.bodyLineEquivalentsUsed.toFixed(2)}/${deterministic.lastPage.bodyLineCapacity} body-line equivalents; ${deterministic.lastPage.visualLines} visual lines`,
    );
  }
  for (const section of deterministic.sections ?? []) {
    const label =
      section.heading === null
        ? "preamble"
        : `H${section.heading.level} ${JSON.stringify(section.heading.title)}`;
    const pages =
      section.pages.length === 0
        ? "0 pages"
        : `pages ${section.pages.map((page) => page.page).join(",")} (${section.pageCount})`;
    const beyond =
      section.pageBudget && !section.pageBudget.withinLimit
        ? `; beyond limit ${section.pageBudget.limitPages}: ${section.pageBudget.pagesBeyondLimit.join(",")}`
        : "";
    rows.push(
      `Section ${section.index} (${label}): ${pages}; ${section.visualLines} visual lines; ${section.countedLines} counted lines${beyond}`,
    );
  }
  if (measurement.renderers.word?.status === "ok") {
    rows.push(
      `Microsoft Word pages: ${measurement.renderers.word.value.pageCount} (delta ${measurement.renderers.word.value.pageCount - deterministic.pageCount})`,
    );
  }
  if (measurement.renderers.libreoffice?.status === "ok") {
    rows.push(
      `LibreOffice Writer pages: ${measurement.renderers.libreoffice.value.pageCount} (delta ${measurement.renderers.libreoffice.value.pageCount - deterministic.pageCount})`,
    );
  }
  for (const paragraph of deterministic.paragraphs ?? []) {
    const filled = Math.round(paragraph.lastLineRatio * 10);
    rows.push(
      `Lines ${paragraph.position.start.line}-${paragraph.position.end.line}: ${Math.round(paragraph.lastLineRatio * 100)}% ${"█".repeat(filled)}${"░".repeat(10 - filled)} ${paragraph.preview}`,
    );
  }
  return `${rows.join("\n")}\n`;
}

function nextSequence(state: SequenceState) {
  return ++state.sequence;
}

function envelope(
  state: SequenceState,
  mode: "single" | "batch" | "watch",
  source: Source,
  measurement: MeasurementResult,
  requestId: string | number | null = null,
  trigger: unknown = null,
) {
  return {
    schemaVersion: 1,
    kind: "result",
    mode,
    sequence: nextSequence(state),
    requestId,
    source,
    trigger,
    measurement: serializableMeasurement(measurement),
  };
}

function errorObject(error: unknown) {
  return error instanceof MdPageCountError
    ? {
        code: error.code,
        message: error.message,
        ...(error.details ? { details: error.details } : {}),
      }
    : {
        code: "INVALID_ARGUMENT",
        message: error instanceof Error ? error.message : String(error),
      };
}

function errorStatus(error: unknown) {
  if (error instanceof MdPageCountError && error.code === "INVALID_ARGUMENT") {
    return 2;
  }
  if (
    error instanceof MdPageCountError &&
    /WORD_|LIBREOFFICE_|NO_OFFICE/.test(error.code)
  ) {
    return 4;
  }
  return 1;
}

async function executeCli(
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
  if (command.mode === "inspect") {
    const result = await inspectDocxTemplate(
      await readFile(resolve(runtime.cwd, command.path)),
    );
    await runtime.writeStdout(
      command.json
        ? `${JSON.stringify(result)}\n`
        : `Template: ${result.package.sha256}\nSections: ${result.sections.length}; selected ${result.selectedSection}\n`,
    );
    return 0;
  }

  if (command.mode === "batch-files" || command.mode === "batch-jsonl") {
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
        throw new MdPageCountError(
          "INVALID_ARGUMENT",
          "JSONL batch requires non-TTY stdin and no positionals",
        );
      }
      const text = await strictUtf8(await runtime.readStdin());
      for (const line of text.split(/\r?\n/)) {
        if (!line.trim()) continue;
        let request: unknown;
        try {
          request = JSON.parse(line);
          if (
            !request ||
            typeof request !== "object" ||
            Array.isArray(request)
          ) {
            throw new Error("request must be an object");
          }
          const record = request as Record<string, unknown>;
          if (
            (typeof record.path === "string") ===
            (typeof record.markdown === "string")
          ) {
            throw new Error("exactly one of path or markdown is required");
          }
          const source: Source =
            typeof record.path === "string"
              ? {
                  kind: "file",
                  path: record.path,
                  resolvedPath: resolve(runtime.cwd, record.path),
                }
              : {
                  kind: "inline",
                  name: typeof record.name === "string" ? record.name : null,
                };
          const markdown =
            typeof record.path === "string"
              ? await strictUtf8(
                  await readFile(resolve(runtime.cwd, record.path)),
                )
              : String(record.markdown);
          const measurement = await measureMarkdown(markdown, base.options);
          await runtime.writeStdout(
            `${JSON.stringify(
              envelope(
                state,
                "batch",
                source,
                measurement,
                (record.id as string | number | null | undefined) ?? null,
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
            `${JSON.stringify({
              schemaVersion: 1,
              kind: "error",
              mode: "batch",
              sequence: nextSequence(state),
              requestId: null,
              source: { kind: "stdin" },
              trigger: null,
              error: errorObject(error),
            })}\n`,
          );
        }
      }
    } else {
      for (const source of batchSources) {
        const resolvedPath = source.resolvedPath;
        try {
          const measurement = await measureMarkdown(
            await strictUtf8(await readFile(resolvedPath)),
            base.options,
          );
          await runtime.writeStdout(
            `${JSON.stringify(
              envelope(state, "batch", source, measurement),
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
            `${JSON.stringify({
              schemaVersion: 1,
              kind: "error",
              mode: "batch",
              sequence: nextSequence(state),
              requestId: null,
              source,
              trigger: null,
              error: errorObject(error),
            })}\n`,
          );
        }
      }
    }
    return failed ? 1 : over ? 3 : 0;
  }

  if (command.mode === "watch") {
    const values = command.values;
    const token = command.path;
    const path = resolve(runtime.cwd, token);
    const source: Source = { kind: "file", path: token, resolvedPath: path };
    const debounce =
      typeof values["debounce-ms"] === "string"
        ? asciiInteger(values["debounce-ms"], "--debounce-ms", 0, 60000)
        : 75;
    let timer: NodeJS.Timeout | undefined;
    let running = false;
    let dirty = false;
    const execute = async (trigger: { kind: string; paths: string[] }) => {
      if (running) {
        dirty = true;
        return;
      }
      running = true;
      try {
        const loaded = await optionsFrom(values, runtime.cwd);
        const measurement = await measureMarkdown(
          await strictUtf8(await readFile(path)),
          loaded.options,
        );
        await runtime.writeStdout(
          values.jsonl
            ? `${JSON.stringify(
                envelope(state, "watch", source, measurement, null, trigger),
              )}\n`
            : `\n[${trigger.kind}]\n${human(measurement)}`,
        );
      } catch (error) {
        const text = `${JSON.stringify({
          schemaVersion: 1,
          kind: "error",
          mode: "watch",
          sequence: nextSequence(state),
          requestId: null,
          source,
          trigger,
          error: errorObject(error),
        })}\n`;
        await (values.jsonl
          ? runtime.writeStdout(text)
          : runtime.writeStderr(text));
      } finally {
        running = false;
        if (dirty) {
          dirty = false;
          void execute({ kind: "dependency-change", paths: [path] });
        }
      }
    };
    const loaded = await optionsFrom(values, runtime.cwd);
    const dependencies = [path, ...loaded.dependencies];
    const watcher = chokidar.watch(dependencies, {
      ignoreInitial: true,
      atomic: 200,
      awaitWriteFinish: { stabilityThreshold: 200, pollInterval: 50 },
      usePolling: values.poll === true,
    });
    if (values.jsonl) {
      await runtime.writeStdout(
        `${JSON.stringify({
          schemaVersion: 1,
          kind: "ready",
          mode: "watch",
          sequence: nextSequence(state),
          source,
          dependencies: dependencies.slice().sort(),
        })}\n`,
      );
    }
    await execute({ kind: "initial", paths: [path] });
    watcher.on("all", (_event, changed) => {
      clearTimeout(timer);
      timer = setTimeout(
        () =>
          void execute({
            kind: changed === path ? "source-change" : "dependency-change",
            paths: [resolve(changed)],
          }),
        debounce,
      );
    });
    const { promise, resolve: finish } = Promise.withResolvers<number>();
    for (const [signal, code] of [
      ["SIGINT", 130],
      ["SIGTERM", 143],
    ] as const) {
      runtime.onceSignal(signal, () => {
        void (async () => {
          await watcher.close();
          if (values.jsonl) {
            await runtime.writeStdout(
              `${JSON.stringify({
                schemaVersion: 1,
                kind: "end",
                mode: "watch",
                sequence: nextSequence(state),
                source,
                reason: signal,
              })}\n`,
            );
          }
          finish(code);
        })();
      });
    }
    return promise;
  }

  const loaded = await optionsFrom(command.values, runtime.cwd);
  const outputPath =
    typeof command.values.output === "string"
      ? command.values.output
      : undefined;
  if (outputPath) loaded.options.includeGeneratedDocx = true;
  let markdown: string;
  if (command.input.kind === "file") {
    markdown = await strictUtf8(
      await readFile(resolve(runtime.cwd, command.input.path)),
    );
  } else {
    if (!command.input.explicit && runtime.stdinIsTTY) {
      throw new MdPageCountError(
        "INVALID_ARGUMENT",
        "No input: provide a file or pipe Markdown on stdin",
      );
    }
    markdown = await strictUtf8(await runtime.readStdin());
  }
  const measurement = await measureMarkdown(markdown, loaded.options);
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
    await runtime.writeStdout(human(measurement));
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
      `${JSON.stringify({ error: errorObject(error) })}\n`,
    );
    return errorStatus(error);
  }
}
