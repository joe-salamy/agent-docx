import { resolve } from "node:path";
import { toErrorPayload } from "../errors.js";
import { builtInProfiles } from "../profiles.js";
import { AgentDocxError } from "../types.js";
import { serializableMeasurement } from "../measurement.js";
import type { MeasurementResult } from "../measurement.js";
import type {
  CliErrorRecord,
  CliFatalRecord,
  CliResultRecord,
  CliSequenceState,
  CliSource,
  CliTrigger,
  CliWatchEndRecord,
  CliWatchReadyRecord,
  Source,
} from "../cli-contract.js";
import { normalizedRelativePath } from "./run-batch.js";

type ProfileCatalogEntry = {
  id: string;
  label: string;
  effectiveDate: string | null;
  sourceUrl: string | null;
  sourceCitation: string;
  requestedFontFamily: string;
  filingPageLimits: Record<string, number>;
};

type ProfileCatalog = {
  schemaVersion: 1;
  profiles: ProfileCatalogEntry[];
};

export function profileCatalog(): ProfileCatalog {
  return {
    schemaVersion: 1,
    profiles: Object.values(builtInProfiles).map((profile) => ({
      id: profile.id,
      label: profile.label,
      effectiveDate: profile.effectiveDate,
      sourceUrl: profile.sourceUrl,
      sourceCitation: profile.sourceCitation,
      requestedFontFamily: profile.requestedFontFamily,
      filingPageLimits: { ...profile.filingPageLimits },
    })),
  };
}

export function humanProfiles(catalog: ProfileCatalog) {
  const rows = ["Built-in profiles:"];
  for (const profile of catalog.profiles) {
    const effective = profile.effectiveDate
      ? `; effective ${profile.effectiveDate}`
      : "";
    rows.push(`${profile.id}: ${profile.label}${effective}`);
    rows.push(`  ${profile.sourceCitation}`);
    const limits = Object.entries(profile.filingPageLimits);
    if (limits.length) {
      rows.push(
        `  Filing page limits: ${limits.map(([kind, pages]) => `${kind} ${pages}`).join(", ")}`,
      );
    }
  }
  return `${rows.join("\n")}\n`;
}

export function human(
  measurement: MeasurementResult,
  display: {
    paragraphs: boolean;
    trim: boolean;
    lines: boolean;
    linesPageFilter?: number[];
  },
) {
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
  if (
    measurement.renderers.word?.status === "ok" &&
    measurement.renderers.word.value.paragraphDiagnostics?.status === "error"
  ) {
    rows.push(
      `Word paragraph diagnostics unavailable: ${measurement.renderers.word.value.paragraphDiagnostics.error.message}`,
    );
  }
  if (measurement.renderers.libreoffice?.status === "ok") {
    rows.push(
      `LibreOffice Writer pages: ${measurement.renderers.libreoffice.value.pageCount} (delta ${measurement.renderers.libreoffice.value.pageCount - deterministic.pageCount})`,
    );
  }
  if (display.paragraphs) {
    for (const paragraph of deterministic.paragraphs ?? []) {
      const filled = Math.max(
        0,
        Math.min(10, Math.round(paragraph.lastLineRatio * 10)),
      );
      rows.push(
        `Lines ${paragraph.position.start.line}-${paragraph.position.end.line}: ${Math.round(paragraph.lastLineRatio * 100)}% ${"█".repeat(filled)}${"░".repeat(10 - filled)} ${paragraph.preview}`,
      );
    }
  }
  if (display.trim) {
    const opportunities = deterministic.trimOpportunities ?? [];
    if (opportunities.length === 0) {
      rows.push("Trim opportunities: none");
    } else {
      rows.push("Trim opportunities:");
      for (const opportunity of opportunities) {
        rows.push(
          `${opportunity.rank}. Lines ${opportunity.position.start.line}-${opportunity.position.end.line}: ${JSON.stringify(opportunity.lastLineText)}; ${Math.round(opportunity.lastLineRatio * 100)}%; ${opportunity.oneLineReduction!.estimatedRemovalTwips} twips; ${opportunity.message}`,
        );
      }
    }
  }
  if (display.lines) {
    const allLines = measurement.deterministic.lines ?? [];
    const lines = display.linesPageFilter
      ? allLines.filter((l) => display.linesPageFilter!.includes(l.page))
      : allLines;
    if (lines.length === 0) {
      rows.push(
        `Lines: none${display.linesPageFilter ? ` (page ${display.linesPageFilter.join(",")})` : ""}`,
      );
    } else {
      rows.push(
        `Lines: ${lines.length} body lines${display.linesPageFilter ? ` (page ${display.linesPageFilter.join(",")})` : ""}`,
      );
      for (const l of lines) {
        const pct = Math.round(l.ratio * 100);
        const filled = Math.max(0, Math.min(10, Math.round(l.ratio * 10)));
        rows.push(
          `P${l.page} L${l.globalIndex + 1} [${l.indexInBlock + 1}/${l.visualLinesInBlock}] ${pct}% ${"█".repeat(filled)}${"░".repeat(10 - filled)} ${l.unusedTwips}twips slack ${l.isLastLineOfBlock ? "(last)" : ""} ${l.text.slice(0, 60).replace(/\s+/g, " ")}`,
        );
      }
    }
  }
  return `${rows.join("\n")}\n`;
}

function nextSequence(state: CliSequenceState) {
  return ++state.sequence;
}

function publicSource(source: Source, cwd: string): CliSource {
  return source.kind === "file"
    ? {
        kind: "file",
        path: normalizedRelativePath(cwd, source.resolvedPath),
      }
    : source;
}

function publicTrigger(
  trigger: CliTrigger | null,
  cwd: string,
): CliTrigger | null {
  return trigger
    ? {
        kind: trigger.kind,
        paths: trigger.paths.map((path) =>
          normalizedRelativePath(cwd, resolve(path)),
        ),
      }
    : null;
}

export function resultRecord(
  state: CliSequenceState,
  mode: "batch" | "watch",
  source: Source,
  measurement: MeasurementResult,
  cwd: string,
  requestId: string | number | null = null,
  trigger: CliTrigger | null = null,
): CliResultRecord {
  return {
    schemaVersion: 1,
    kind: "result",
    mode,
    sequence: nextSequence(state),
    requestId,
    source: publicSource(source, cwd),
    trigger: publicTrigger(trigger, cwd),
    measurement: serializableMeasurement(measurement),
  };
}

export function errorRecord(
  state: CliSequenceState,
  mode: "batch" | "watch",
  source: Source,
  error: unknown,
  cwd: string,
  requestId: string | number | null = null,
  trigger: CliTrigger | null = null,
): CliErrorRecord {
  return {
    schemaVersion: 1,
    kind: "error",
    mode,
    sequence: nextSequence(state),
    requestId,
    source: publicSource(source, cwd),
    trigger: publicTrigger(trigger, cwd),
    error: toErrorPayload(error),
  };
}

export function readyRecord(
  state: CliSequenceState,
  source: Source,
  dependencies: readonly string[],
  cwd: string,
): CliWatchReadyRecord {
  return {
    schemaVersion: 1,
    kind: "ready",
    mode: "watch",
    sequence: nextSequence(state),
    source: publicSource(source, cwd),
    dependencies: dependencies
      .map((path) => normalizedRelativePath(cwd, resolve(path)))
      .sort(),
  };
}
export function endRecord(
  state: CliSequenceState,
  source: Source,
  reason: "SIGINT" | "SIGTERM",
  cwd: string,
): CliWatchEndRecord {
  return {
    schemaVersion: 1,
    kind: "end",
    mode: "watch",
    sequence: nextSequence(state),
    source: publicSource(source, cwd),
    reason,
  };
}
export function fatalRecord(error: unknown): CliFatalRecord {
  return { schemaVersion: 1, kind: "fatal", error: toErrorPayload(error) };
}
export function agentFatalRecord(error: unknown, sequence: number) {
  return {
    schemaVersion: 1,
    kind: "fatal" as const,
    sequence,
    requestId: null,
    action: null,
    project: null,
    documentId: null,
    revision: null,
    error: toErrorPayload(error),
  };
}

export function errorStatus(error: unknown) {
  if (error instanceof AgentDocxError && error.code === "INVALID_ARGUMENT") {
    return 2;
  }
  if (
    error instanceof AgentDocxError &&
    /WORD_|LIBREOFFICE_|NO_OFFICE/.test(error.code)
  ) {
    return 4;
  }
  return 1;
}
