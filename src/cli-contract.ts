import type { ErrorCode, JsonValue } from "./types.js";
import type {
  EstimateOptions,
  MeasurementResult,
  RendererMode,
} from "./measurement.js";
import type { LayoutOverrides } from "./layout/profile.js";

export type CliJsonlRequest =
  | { id?: string | number | null; path: string }
  | { id?: string | number | null; name?: string; markdown: string };

export type CliSource =
  | { kind: "file"; path: string }
  | { kind: "stdin" }
  | { kind: "inline"; name: string | null };

export type CliTrigger = {
  kind: "initial" | "source-change" | "dependency-change";
  paths: readonly string[];
};

export type CliErrorPayload = {
  code: ErrorCode;
  message: string;
  details?: Record<string, JsonValue>;
};

export type CliResultRecord = {
  schemaVersion: 1;
  kind: "result";
  mode: "batch" | "watch";
  sequence: number;
  requestId: string | number | null;
  source: CliSource;
  trigger: CliTrigger | null;
  measurement: MeasurementResult;
};

export type CliErrorRecord = {
  schemaVersion: 1;
  kind: "error";
  mode: "batch" | "watch";
  sequence: number;
  requestId: string | number | null;
  source: CliSource;
  trigger: CliTrigger | null;
  error: CliErrorPayload;
};

export type CliWatchReadyRecord = {
  schemaVersion: 1;
  kind: "ready";
  mode: "watch";
  sequence: number;
  source: CliSource;
  dependencies: readonly string[];
};

export type CliWatchEndRecord = {
  schemaVersion: 1;
  kind: "end";
  mode: "watch";
  sequence: number;
  source: CliSource;
  reason: "SIGINT" | "SIGTERM";
};

export type CliFatalRecord = {
  schemaVersion: 1;
  kind: "fatal";
  error: CliErrorPayload;
};

export interface CliRuntime {
  readonly cwd: string;
  readonly stdinIsTTY: boolean;
  readonly version: string;
  readStdin(): Promise<Uint8Array>;
  readStdinChunks(): AsyncIterable<Uint8Array>;
  writeStdout(text: string): Promise<void>;
  writeStderr(text: string): Promise<void>;
  onceSignal(signal: "SIGINT" | "SIGTERM", listener: () => void): void;
}

export type Source =
  | { kind: "file"; path: string; resolvedPath: string }
  | { kind: "stdin" }
  | { kind: "inline"; name: string | null };

export type BatchSelection = {
  readonly recursive: boolean;
  readonly include: readonly string[];
  readonly exclude: readonly string[];
};

export type CliSequenceState = { sequence: number };

export type OutputFileHandle = {
  writeFile(bytes: Uint8Array): Promise<void>;
  sync?(): Promise<void>;
  close(): Promise<void>;
};
export type OutputFileIo = {
  open(path: string, flags: "wx"): Promise<OutputFileHandle>;
  link?(existingPath: string, newPath: string): Promise<void>;
  unlink(path: string): Promise<void>;
};

export type SerializableConfig = {
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
  lineDiagnostics?: boolean;
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

export type { MeasurementResult };
