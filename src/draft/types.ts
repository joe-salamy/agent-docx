import type { MeasurementResult } from "../measurement.js";
import type { ErrorCode, JsonValue } from "../types.js";
import type { BlockId, RevisionId } from "../legal/model.js";
import type { ValidationResult } from "../legal/rules.js";

export type DraftEvaluationError = {
  code: ErrorCode;
  message: string;
  details?: Record<string, JsonValue>;
};

export type SourcePatch = {
  schemaVersion: 1;
  documentId: string;
  baseRevision: RevisionId;
  edits: readonly {
    start: number;
    end: number;
    expectedText: string;
    replacement: string;
  }[];
};

export type PatchDeltas = {
  pageCount: number;
  countedLines: number;
  pageLimitExcess: number;
  countedLineExcess: number;
  lastPageUsedTwips: number;
  validationSummary: { pass: number; fail: number; unknown: number };
  affected: readonly {
    blockId: BlockId;
    sourceRanges: readonly { start: number; end: number }[];
    pageSpan: number;
    lineCount: number;
    lastLineUsedTwips: number;
  }[];
  sections: readonly {
    sectionIndex: number;
    pageCount: number;
    countedLines: number;
    usedTwips: number;
  }[];
};

export type PatchEvaluation = {
  schemaVersion: 1;
  documentId: string;
  patchHash: RevisionId;
  baseRevision: RevisionId;
  before: {
    measurement: Omit<MeasurementResult, "generatedDocx">;
    validation: ValidationResult;
  };
  candidate:
    | {
        status: "ok";
        measurement: Omit<MeasurementResult, "generatedDocx">;
        validation: ValidationResult;
        deltas: PatchDeltas;
      }
    | { status: "invalid"; error: DraftEvaluationError };
  passesConstraints: boolean;
  canApply: boolean;
  state: {
    headMatchesBase: boolean;
    sourceMatchesBase: boolean;
    documentConfigMatchesBase: boolean;
    dependenciesMatchBase: boolean;
    baseWorkingTreeHash: RevisionId;
    workingTreeHash: RevisionId;
  };
};

export type DraftGuidance = {
  schemaVersion: 1;
  documentId: string;
  revision: RevisionId;
  baseRevision: RevisionId;
  workingTreeHash: RevisionId;
  items: readonly {
    blockId: BlockId;
    pages: readonly number[];
    sectionIndex: number;
    overflowingPage: boolean;
    overflowingSection: boolean;
    oneLineReduction: boolean;
    minimumReduction: { twips: number; lines: number };
    editableSourceRanges: readonly { start: number; end: number }[];
    lastLine: {
      usedTwips: number;
      availableTwips: number;
      remainingTwips: number;
    };
    budgets: {
      pageLimit: number | null;
      pagesRemaining: number | null;
      countedLineLimit: number | null;
      countedLinesRemaining: number | null;
    };
  }[];
};

