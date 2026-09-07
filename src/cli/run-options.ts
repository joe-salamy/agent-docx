import { readFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import type { CliOptionValues } from "../cli-args.js";
import { MAX_FONT_BYTES, readInputFile } from "../input.js";
import { strictUtf8 } from "../jsonl.js";
import { AgentDocxError } from "../types.js";
import type {
  EstimateOptions,
  MeasureOptions,
  RendererMode,
} from "../measurement.js";
import type { FontSetInput, LayoutOverrides } from "../layout/profile.js";
import type { SerializableConfig } from "../cli-contract.js";

export function asciiInteger(
  value: string,
  name: string,
  min = 1,
  max = Number.MAX_SAFE_INTEGER,
) {
  if (!/^(?:0|[1-9][0-9]*)$/.test(value)) {
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `${name} requires ASCII digits`,
    );
  }
  const number = Number(value);
  if (!Number.isSafeInteger(number) || number < min || number > max) {
    throw new AgentDocxError("INVALID_ARGUMENT", `${name} is out of range`);
  }
  return number;
}

function decimal(value: string, name: string, positive = true) {
  if (!/^(?:0|[1-9][0-9]*)(?:[.][0-9]+)?$/.test(value)) {
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `${name} has invalid decimal syntax`,
    );
  }
  const number = Number(value);
  if (!Number.isFinite(number) || (positive ? number <= 0 : number < 0)) {
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `${name} must be ${positive ? "positive" : "nonnegative"}`,
    );
  }
  return number;
}

const twips = (number: number, scale: number) =>
  Math.floor(number * scale + 0.5);

async function loadConfig(pathToken: string): Promise<{
  config: SerializableConfig;
  base: string;
  path: string;
}> {
  const path = resolve(pathToken);
  const bytes = await readInputFile(path, "Configuration");
  let value: unknown;
  try {
    value = JSON.parse(await strictUtf8(bytes));
  } catch (error) {
    if (error instanceof AgentDocxError) throw error;
    throw new AgentDocxError(
      "INVALID_CONFIG",
      "Configuration is not valid JSON",
    );
  }
  const schemaPath = fileURLToPath(
    new URL("../../schemas/config.schema.json", import.meta.url),
  );
  const schema = JSON.parse(await readFile(schemaPath, "utf8")) as object;
  // Lazy Ajv: only needed when --config is supplied; avoids 5s 9p load for --help/profiles
  const { Ajv2020 } = await import("ajv/dist/2020.js");
  const ajv = new Ajv2020({ allErrors: true, strict: true });
  const validate = ajv.compile(schema);
  if (!validate(value)) {
    throw new AgentDocxError(
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
  return readInputFile(resolve(base, path), "Font", MAX_FONT_BYTES);
}

export async function optionsFrom(
  values: CliOptionValues,
  cwd: string,
): Promise<{
  options: MeasureOptions;
  dependencies: string[];
  batch: SerializableConfig["batch"];
  linePageFilter?: number[];
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
  if ((config as Record<string, unknown>).lineDiagnostics === true) {
    options.lineDiagnostics = true;
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
    // Lazy docx inspect: yauzl/saxes heavy; only for --template
    const { inspectDocxTemplate } = await import("../docx/inspect.js");
    options.template = await inspectDocxTemplate(
      await readInputFile(templateToken, "DOCX template"),
      {
        ...(options.profile !== undefined
          ? { fallbackProfile: options.profile }
          : {}),
      },
    );
    dependencies.push(templateToken);
  }

  const hasCliFontOption = [
    "font-family",
    "font-regular",
    "font-bold",
    "font-italic",
    "font-bold-italic",
  ].some((key) => values[key] !== undefined);
  if (hasCliFontOption && typeof values["font-regular"] !== "string") {
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      "--font-regular is required with --font-* options",
    );
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
      throw new AgentDocxError(
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
    options.profile = values.profile as NonNullable<MeasureOptions["profile"]>;
  }
  if (typeof values["filing-kind"] === "string") {
    options.filingKind = values["filing-kind"] as NonNullable<
      EstimateOptions["filingKind"]
    >;
  }
  if (typeof values["page-limit"] === "string") {
    options.pageLimit = asciiInteger(values["page-limit"], "--page-limit");
  }
  if (values.paragraphs === true) options.paragraphDiagnostics = true;
  if (values.sections === true) options.sectionDiagnostics = true;
  if (values.lines === true) options.lineDiagnostics = true;
  let linePageFilter: number[] | undefined;
  if (values["lines-page"] !== undefined) {
    if (values.lines !== true) {
      throw new AgentDocxError(
        "INVALID_ARGUMENT",
        "--lines-page requires --lines",
      );
    }
    const raw = values["lines-page"] as readonly string[];
    const tokens: string[] = [];
    for (const entry of raw) {
      for (const part of String(entry).split(",")) {
        const trimmed = part.trim();
        if (trimmed.length) tokens.push(trimmed);
      }
    }
    if (tokens.length === 0) {
      throw new AgentDocxError(
        "INVALID_ARGUMENT",
        "--lines-page requires a page number",
      );
    }
    linePageFilter = tokens.map((t) => asciiInteger(t, "--lines-page", 1));
  }
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
    throw new AgentDocxError(
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
  return {
    options,
    dependencies,
    batch: config.batch,
    ...(linePageFilter ? { linePageFilter } : {}),
  };
}
