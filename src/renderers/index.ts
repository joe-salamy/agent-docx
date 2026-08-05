import { estimateNormalizedDocument } from "../estimate.js";
import { normalizeMarkdown, type NormalizedDocument } from "../markdown.js";
import { renderLibreOffice, renderWord } from "./office.js";
import { toErrorPayload } from "../errors.js";
import { AgentDocxError } from "../types.js";
import type {
  Budget,
  LibreOfficeRendering,
  MeasureOptions,
  MeasurementResult,
  PageCountSource,
  RendererError,
  RendererMode,
  RendererStatus,
  WordRendering,
} from "../measurement.js";

function rendererError(error: unknown, phase: string): RendererError {
  const projected = toErrorPayload(error);
  return {
    ...projected,
    phase,
  };
}
function budgetFor(
  limit: number,
  pageCount: number,
  equivalent: number,
  linesRemaining: number,
): Budget {
  return {
    limitPages: limit,
    withinLimit: pageCount <= limit,
    pagesRemaining: limit - pageCount,
    equivalentPagesRemaining: limit - equivalent,
    bodyLineEquivalentsRemaining: linesRemaining,
    fractionalFieldsSource: "deterministic",
  };
}
function validateRenderer(mode: string): RendererMode {
  if (
    !(
      ["deterministic", "word", "libreoffice", "compare"] as readonly string[]
    ).includes(mode)
  )
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `Unknown renderer mode: ${mode}`,
    );
  return mode as RendererMode;
}

export async function measureNormalizedDocument(
  flow: NormalizedDocument,
  options: MeasureOptions = {},
): Promise<MeasurementResult> {
  const mode = validateRenderer(options.renderer ?? "deterministic");
  const timeout = options.officeTimeoutMs;
  if (
    timeout !== undefined &&
    (!Number.isInteger(timeout) || timeout < 1000 || timeout > 600000)
  )
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      "officeTimeoutMs must be an integer from 1000 through 600000",
    );
  const deterministic = await estimateNormalizedDocument(flow, options);
  const output: MeasurementResult = {
    schemaVersion: 1,
    mode,
    pageCount: deterministic.pageCount,
    pageCountSource: "deterministic",
    deterministic,
    renderers: {},
  };
  const pagination = deterministic.profile.pagination;
  if (
    mode !== "deterministic" &&
    pagination.widowOrphanControl &&
    (pagination.orphanLines !== 2 || pagination.widowLines !== 2)
  )
    deterministic.warnings = [
      ...deterministic.warnings,
      {
        code: "RENDERER_PARITY_APPROXIMATE",
        severity: "warning",
        message:
          "Word expresses widow/orphan control as on/off (2 lines); configured orphanLines/widowLines differ, so Office pagination may diverge.",
      },
    ];
  if (deterministic.budget) output.budget = deterministic.budget;
  const needsDocx =
    options.includeGeneratedDocx === true || mode !== "deterministic";
  if (!needsDocx) return output;
  // Load DOCX generation only when bytes or Office rendering are requested.
  const { generateDocx } = await import("../docx/generate.js");
  const { loadFonts } = await import("../resolve.js");
  const fonts = await loadFonts(
    options.fontSet,
    deterministic.profile.requestedFontFamily,
  );
  const generated = await generateDocx(flow, deterministic.profile, {
    ...(options.chrome !== undefined ? { chrome: options.chrome } : {}),
    ...(options.metadata !== undefined ? { metadata: options.metadata } : {}),
    ...(options.assets !== undefined ? { assets: options.assets } : {}),
    pageCount: deterministic.pageCount,
    fonts,
  });
  const docx = generated.bytes;
  if (options.includeGeneratedDocx === true) output.generatedDocx = docx;
  if (mode === "deterministic") return output;
  const requestedFontFamilies = [deterministic.profile.requestedFontFamily];
  let word: RendererStatus<WordRendering> | undefined;
  let libreoffice: RendererStatus<LibreOfficeRendering> | undefined;
  if (mode === "word" || mode === "compare") {
    try {
      word = {
        status: "ok",
        value: await renderWord(
          docx,
          requestedFontFamilies,
          options.word,
          timeout ?? 120000,
          options.paragraphDiagnostics || options.trim
            ? generated.bodyParagraphs
            : undefined,
        ),
      };
    } catch (error) {
      word = {
        status: /NOT_FOUND|UNAVAILABLE/.test(
          error instanceof AgentDocxError ? error.code : "",
        )
          ? "unavailable"
          : "error",
        error: rendererError(error, "word"),
      };
      if (mode === "word") throw error;
    }
  }
  if (mode === "libreoffice" || mode === "compare") {
    try {
      libreoffice = {
        status: "ok",
        value: await renderLibreOffice(
          docx,
          requestedFontFamilies,
          options.libreoffice,
          timeout ?? 60000,
        ),
      };
    } catch (error) {
      libreoffice = {
        status: /NOT_FOUND|UNAVAILABLE/.test(
          error instanceof AgentDocxError ? error.code : "",
        )
          ? "unavailable"
          : "error",
        error: rendererError(error, "libreoffice"),
      };
      if (mode === "libreoffice") throw error;
    }
  }
  if (word) output.renderers.word = word;
  if (libreoffice) output.renderers.libreoffice = libreoffice;
  if (
    libreoffice?.status === "ok" &&
    !libreoffice.value.calibratedFontEnvironment
  )
    deterministic.warnings = [
      ...deterministic.warnings,
      {
        code: "LIBREOFFICE_FONT_ENVIRONMENT_UNVERIFIED",
        severity: "warning",
        message:
          "LibreOffice font files were not supplied for verifiable renderer font parity.",
      },
    ];
  if (mode === "word" && word?.status === "ok") {
    output.pageCount = word.value.pageCount;
    output.pageCountSource = "word";
  } else if (mode === "libreoffice" && libreoffice?.status === "ok") {
    output.pageCount = libreoffice.value.pageCount;
    output.pageCountSource = "libreoffice";
  } else if (
    mode === "compare" &&
    word?.status !== "ok" &&
    libreoffice?.status !== "ok"
  )
    throw new AgentDocxError(
      "NO_OFFICE_RENDERER",
      "Neither Microsoft Word nor LibreOffice rendered successfully",
    );
  const limit = deterministic.budget?.limitPages;
  if (limit !== undefined) {
    const equivalent = deterministic.equivalentPages;
    const lines = deterministic.budget!.bodyLineEquivalentsRemaining;
    const budgets: Partial<Record<PageCountSource, Budget>> = {
      deterministic: deterministic.budget!,
    };
    if (word?.status === "ok")
      budgets.word = budgetFor(limit, word.value.pageCount, equivalent, lines);
    if (libreoffice?.status === "ok")
      budgets.libreoffice = budgetFor(
        limit,
        libreoffice.value.pageCount,
        equivalent,
        lines,
      );
    output.budgetBySource = budgets;
    // pageCountSource always names a key above (deterministic, or the ok renderer).
    output.budget = budgets[output.pageCountSource]!;
  }
  return output;
}

export async function measureMarkdown(
  markdown: string,
  options: MeasureOptions = {},
): Promise<MeasurementResult> {
  if (typeof markdown !== "string")
    throw new AgentDocxError("INVALID_ARGUMENT", "markdown must be a string");
  return measureNormalizedDocument(normalizeMarkdown(markdown), options);
}
