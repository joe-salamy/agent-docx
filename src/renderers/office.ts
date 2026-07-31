import {
  access,
  mkdtemp,
  mkdir,
  readFile,
  rm,
  stat,
  writeFile,
} from "node:fs/promises";
import { constants } from "node:fs";
import { tmpdir } from "node:os";
import { delimiter, isAbsolute, join } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { spawn } from "node:child_process";
import { createHash } from "node:crypto";
import { PDFDocument } from "pdf-lib";
import {
  AgentDocxError,
  type LibreOfficeRendererOptions,
  type LibreOfficeRendering,
  type WordRendererOptions,
  type RendererError,
  type WordParagraphDiagnostic,
  type WordRendering,
} from "../types.js";
import type { BodyParagraphManifestEntry } from "../docx/generate.js";
const sha = (b: Uint8Array) => createHash("sha256").update(b).digest("hex");
type ProcessResult = {
  code: number | null;
  stdout: string;
  stderr: string;
  timedOut: boolean;
  stdoutOverflow: boolean;
  stderrOverflow: boolean;
};
async function run(
  executable: string,
  args: string[],
  stdin: string | undefined,
  timeoutMs: number,
  env?: NodeJS.ProcessEnv,
): Promise<ProcessResult> {
  const child = spawn(executable, args, {
    shell: false,
    stdio: ["pipe", "pipe", "pipe"],
    detached: process.platform !== "win32",
    env,
  });
  let stopped = false;
  const terminate = () => {
    if (stopped) return;
    stopped = true;
    try {
      if (process.platform !== "win32" && child.pid)
        process.kill(-child.pid, "SIGKILL");
      else child.kill("SIGKILL");
    } catch {
      try {
        child.kill("SIGKILL");
      } catch {}
    }
  };
  const chunksOut: Buffer[] = [],
    chunksErr: Buffer[] = [];
  let out = 0,
    err = 0,
    timedOut = false;
  const stdoutLimit = 4 * 1024 * 1024;
  const stderrLimit = 1024 * 1024;
  child.stdout.on("data", (buffer: Buffer) => {
    const remaining = Math.max(0, stdoutLimit - out);
    if (remaining > 0) chunksOut.push(buffer.subarray(0, remaining));
    out += buffer.length;
    if (out > stdoutLimit) terminate();
  });
  child.stderr.on("data", (buffer: Buffer) => {
    const remaining = Math.max(0, stderrLimit - err);
    if (remaining > 0) chunksErr.push(buffer.subarray(0, remaining));
    err += buffer.length;
    if (err > stderrLimit) terminate();
  });
  if (stdin === undefined) child.stdin.end();
  else child.stdin.end(stdin);
  const timer = setTimeout(() => {
    timedOut = true;
    terminate();
  }, timeoutMs);
  const { promise, resolve, reject } = Promise.withResolvers<number | null>();
  child.once("error", reject);
  child.once("close", resolve);
  let code: number | null;
  try {
    code = await promise;
  } finally {
    clearTimeout(timer);
  }
  return {
    code,
    stdout: Buffer.concat(chunksOut).toString("utf8"),
    stderr: Buffer.concat(chunksErr).toString("utf8"),
    timedOut,
    stdoutOverflow: out > stdoutLimit,
    stderrOverflow: err > stderrLimit,
  };
}
async function exists(path: string) {
  try {
    await access(path, constants.X_OK);
    return true;
  } catch {
    return false;
  }
}
export async function resolveLibreOffice(explicit?: string): Promise<string> {
  if (explicit) {
    if (!isAbsolute(explicit))
      throw new AgentDocxError(
        "INVALID_ARGUMENT",
        "LibreOffice executable path must be absolute",
      );
    if (await exists(explicit)) return explicit;
    throw new AgentDocxError(
      "LIBREOFFICE_NOT_FOUND",
      `LibreOffice executable not found: ${explicit}`,
    );
  }
  const candidates: string[] = [];
  if (process.platform === "darwin")
    candidates.push("/Applications/LibreOffice.app/Contents/MacOS/soffice");
  if (process.platform === "win32") {
    for (const base of [
      process.env.ProgramFiles,
      process.env["ProgramFiles(x86)"],
    ])
      if (base)
        candidates.push(join(base, "LibreOffice", "program", "soffice.com"));
  } else
    for (const dir of (process.env.PATH ?? "").split(delimiter))
      for (const name of ["soffice", "libreoffice"])
        candidates.push(join(dir, name));
  for (const candidate of candidates)
    if (await exists(candidate)) return candidate;
  throw new AgentDocxError(
    "LIBREOFFICE_NOT_FOUND",
    "LibreOffice Writer executable was not found",
  );
}
export async function renderLibreOffice(
  docx: Uint8Array,
  requestedFontFamilies: readonly string[],
  options: LibreOfficeRendererOptions = {},
  timeoutMs = 60000,
): Promise<LibreOfficeRendering> {
  const executable = await resolveLibreOffice(options.executablePath);
  const versionResult = await run(executable, ["--version"], undefined, 10000);
  if (versionResult.stdoutOverflow || versionResult.stderrOverflow)
    throw new AgentDocxError(
      "LIBREOFFICE_RENDER_FAILED",
      "LibreOffice version output exceeded the transport limit",
    );
  if (versionResult.code !== 0)
    throw new AgentDocxError(
      "LIBREOFFICE_RENDER_FAILED",
      "LibreOffice version preflight failed",
      { stderr: versionResult.stderr },
    );
  const root = await mkdtemp(join(tmpdir(), "agent-docx-lo-"));
  const inputDir = join(root, "input"),
    outputDir = join(root, "output"),
    profileDir = join(root, "profile");
  const input = join(inputDir, "render.docx");
  const start = performance.now();
  try {
    await Promise.all([mkdir(inputDir), mkdir(outputDir), mkdir(profileDir)]);
    await writeFile(input, docx);
    await writeFile(
      join(profileDir, "registrymodifications.xcu"),
      '<?xml version="1.0" encoding="UTF-8"?><oor:items xmlns:oor="http://openoffice.org/2001/registry"></oor:items>',
    );
    const args = [
      `-env:UserInstallation=${pathToFileURL(profileDir).href}`,
      "--headless",
      "--nologo",
      "--nodefault",
      "--norestore",
      "--infilter=Office Open XML Text",
      "--convert-to",
      "pdf:writer_pdf_Export",
      "--outdir",
      outputDir,
      input,
    ];
    const rendered = await run(executable, args, undefined, timeoutMs, {
      ...process.env,
      LC_ALL: "C",
      LANG: "C",
      TZ: "UTC",
    });
    if (rendered.stdoutOverflow || rendered.stderrOverflow)
      throw new AgentDocxError(
        "LIBREOFFICE_RENDER_FAILED",
        "LibreOffice output exceeded the transport limit",
      );
    if (rendered.timedOut)
      throw new AgentDocxError(
        "LIBREOFFICE_TIMEOUT",
        `LibreOffice exceeded ${timeoutMs} ms`,
      );
    if (rendered.code !== 0)
      throw new AgentDocxError(
        "LIBREOFFICE_RENDER_FAILED",
        "LibreOffice conversion failed",
        { stderr: rendered.stderr },
      );
    const pdfPath = join(outputDir, "render.pdf");
    const info = await stat(pdfPath);
    if (!info.isFile() || info.isSymbolicLink() || info.size === 0)
      throw new AgentDocxError(
        "LIBREOFFICE_RENDER_FAILED",
        "LibreOffice did not create a regular nonempty PDF",
      );
    const pdf = await readFile(pdfPath);
    let loaded: PDFDocument;
    try {
      loaded = await PDFDocument.load(pdf, {
        ignoreEncryption: false,
        updateMetadata: false,
        throwOnInvalidObject: true,
      });
    } catch {
      throw new AgentDocxError(
        "LIBREOFFICE_RENDER_FAILED",
        "LibreOffice produced an invalid or encrypted PDF",
      );
    }
    const pageCount = loaded.getPageCount();
    if (pageCount < 1)
      throw new AgentDocxError(
        "LIBREOFFICE_RENDER_FAILED",
        "LibreOffice PDF has no pages",
      );
    const calibrated = Boolean(options.installedFonts?.length);
    return {
      pageCount,
      versionRaw: versionResult.stdout.trim() || versionResult.stderr.trim(),
      executablePath: executable,
      platform: process.platform,
      arch: process.arch,
      calibratedFontEnvironment: calibrated,
      requestedFontFamilies,
      generatedDocxSha256: sha(docx),
      pdfSha256: sha(pdf),
      durationMs: performance.now() - start,
    };
  } finally {
    await rm(root, {
      recursive: true,
      force: true,
      maxRetries: 4,
      retryDelay: 50,
    });
  }
}
async function resolveWordPath(options: WordRendererOptions) {
  if (options.powerShellPath) {
    if (!isAbsolute(options.powerShellPath))
      throw new AgentDocxError(
        "INVALID_ARGUMENT",
        "PowerShell path must be absolute",
      );
    if (await exists(options.powerShellPath)) return options.powerShellPath;
    throw new AgentDocxError(
      "WORD_NOT_FOUND",
      `PowerShell not found: ${options.powerShellPath}`,
    );
  }
  if (process.platform === "win32") {
    const root = process.env.SystemRoot ?? "C:\\Windows";
    const candidate = join(
      root,
      "System32",
      "WindowsPowerShell",
      "v1.0",
      "powershell.exe",
    );
    if (await exists(candidate)) return candidate;
  }
  if (
    /microsoft/i.test(
      await readFile("/proc/sys/kernel/osrelease", "utf8").catch(() => ""),
    )
  ) {
    const candidate =
      "/mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe";
    if (await exists(candidate)) return candidate;
  }
  throw new AgentDocxError(
    "WORD_NOT_FOUND",
    "Microsoft Word PowerShell bridge is unavailable",
  );
}
export async function renderWord(
  docx: Uint8Array,
  requestedFontFamilies: readonly string[],
  options: WordRendererOptions = {},
  timeoutMs = 120000,
  paragraphManifest?: readonly BodyParagraphManifestEntry[],
): Promise<WordRendering> {
  if (docx.byteLength > 25 * 1024 * 1024)
    throw new AgentDocxError(
      "WORD_RENDER_FAILED",
      "Generated DOCX exceeds Word adapter limit",
    );
  const powershell = await resolveWordPath(options);
  let script = fileURLToPath(
    new URL("../../assets/word/render.ps1", import.meta.url),
  );
  if (powershell.startsWith("/mnt/") && process.platform !== "win32") {
    if (!(await exists("/usr/bin/wslpath")))
      throw new AgentDocxError(
        "WORD_WSL_BRIDGE_UNAVAILABLE",
        "/usr/bin/wslpath is unavailable",
      );
    const converted = await run(
      "/usr/bin/wslpath",
      ["-w", script],
      undefined,
      5000,
    );
    if (
      converted.code !== 0 ||
      converted.stdoutOverflow ||
      converted.stderrOverflow
    )
      throw new AgentDocxError(
        "WORD_WSL_BRIDGE_UNAVAILABLE",
        "wslpath failed",
      );
    script = converted.stdout.trim();
  }
  const start = performance.now();
  const paragraphIds = paragraphManifest?.map((entry) => entry.id) ?? [];
  const response = await run(
    powershell,
    ["-NoLogo", "-NoProfile", "-NonInteractive", "-File", script],
    JSON.stringify({
      protocolVersion: 2,
      docxBase64: Buffer.from(docx).toString("base64"),
      requestedFontFamilies,
      paragraphIds,
    }),
    timeoutMs,
  );
  if (response.timedOut)
    throw new AgentDocxError("WORD_TIMEOUT", `Word exceeded ${timeoutMs} ms`);
  const frames = response.stdout
    .split(/\r?\n/)
    .filter(Boolean)
    .map((line) => {
      try {
        return JSON.parse(line) as unknown;
      } catch {
        return null;
      }
    });
  const expectedKinds = paragraphIds.length
    ? ["started", "summary", "paragraphs"]
    : ["started", "summary"];
  const first = frames[0] as Record<string, unknown> | null | undefined;
  const validSequence =
    frames.length === expectedKinds.length &&
    frames.every(
      (frame, index) =>
        frame !== null &&
        typeof frame === "object" &&
        !Array.isArray(frame) &&
        (frame as Record<string, unknown>).kind === expectedKinds[index],
    ) &&
    first?.protocolVersion === 2;
  const summaries = frames.filter(
    (frame) =>
      frame !== null &&
      typeof frame === "object" &&
      "kind" in frame &&
      frame.kind === "summary",
  );
  if (response.code !== 0 || !validSequence || summaries.length !== 1)
    throw new AgentDocxError(
      "WORD_RENDER_FAILED",
      "Word adapter did not return exactly one successful summary",
      { stderr: response.stderr, stdout: response.stdout },
    );
  const summary = summaries[0] as Record<string, unknown>;
  const pageCount = summary.pageCount;
  const totalBodyLines = summary.totalBodyLines;
  const bodyLinesByPage = summary.bodyLinesByPage;
  const bodyLinesOnLastPage = summary.bodyLinesOnLastPage;
  if (
    summary.protocolVersion !== 2 ||
    !Number.isInteger(pageCount) ||
    (pageCount as number) < 1 ||
    !Number.isInteger(totalBodyLines) ||
    (totalBodyLines as number) < 0 ||
    !Array.isArray(bodyLinesByPage) ||
    bodyLinesByPage.length !== pageCount ||
    !bodyLinesByPage.every(
      (lineCount) => Number.isInteger(lineCount) && lineCount >= 0,
    ) ||
    (bodyLinesOnLastPage !== null &&
      (!Number.isInteger(bodyLinesOnLastPage) ||
        (bodyLinesOnLastPage as number) < 0)) ||
    typeof summary.version !== "string" ||
    typeof summary.build !== "string" ||
    typeof summary.activePrinter !== "string"
  )
    throw new AgentDocxError(
      "WORD_RENDER_FAILED",
      "Word adapter returned a malformed version-2 summary",
      { stdout: response.stdout },
    );
  let paragraphDiagnostics: WordRendering["paragraphDiagnostics"];
  if (paragraphManifest?.length) {
    let paragraphError: RendererError | undefined;
    const failParagraphs = (message: string): RendererError => ({
      code: "WORD_RENDER_FAILED",
      message,
      phase: "paragraph-diagnostics",
    });
    if (response.stdoutOverflow || response.stderrOverflow) {
      paragraphError = failParagraphs(
        "Word paragraph diagnostic output exceeded the transport limit",
      );
    }
    const paragraphFrames = frames.filter(
      (frame) =>
        frame !== null &&
        typeof frame === "object" &&
        "kind" in frame &&
        frame.kind === "paragraphs",
    );
    if (!paragraphError && paragraphFrames.length !== 1)
      paragraphError = failParagraphs(
        "Word adapter did not return exactly one paragraph frame",
      );
    const paragraphFrame = paragraphFrames[0] as
      | Record<string, unknown>
      | undefined;
    if (
      !paragraphError &&
      (paragraphFrame?.protocolVersion !== 2 ||
        !["ok", "error"].includes(String(paragraphFrame.status)))
    )
      paragraphError = failParagraphs(
        "Word adapter returned a malformed paragraph frame",
      );
    if (!paragraphError && paragraphFrame?.status === "error")
      paragraphError = failParagraphs(
        typeof paragraphFrame.message === "string"
          ? paragraphFrame.message
          : "Word paragraph extraction failed",
      );
    const nativeValues =
      !paragraphError &&
      paragraphFrame?.status === "ok" &&
      Array.isArray(paragraphFrame.value)
        ? paragraphFrame.value
        : undefined;
    if (!paragraphError && !nativeValues)
      paragraphError = failParagraphs(
        "Word adapter returned invalid paragraph values",
      );
    const byId = new Map<string, Record<string, unknown>>();
    if (nativeValues) {
      for (const native of nativeValues) {
        if (
          !native ||
          typeof native !== "object" ||
          Array.isArray(native) ||
          typeof (native as Record<string, unknown>).id !== "string"
        ) {
          paragraphError = failParagraphs(
            "Word adapter returned a malformed paragraph record",
          );
          break;
        }
        const value = native as Record<string, unknown>;
        const id = value.id as string;
        if (byId.has(id)) {
          paragraphError = failParagraphs(
            "Word adapter returned duplicate paragraph IDs",
          );
          break;
        }
        byId.set(id, value);
      }
    }
    const exactIds =
      byId.size === paragraphIds.length &&
      paragraphIds.every((id) => byId.has(id));
    if (!paragraphError && !exactIds)
      paragraphError = failParagraphs(
        "Word adapter paragraph IDs did not match the request",
      );
    const mapped: WordParagraphDiagnostic[] = [];
    if (!paragraphError) {
      for (const manifest of paragraphManifest) {
        const native = byId.get(manifest.id)!;
        if (
          !Number.isInteger(native.lineCount) ||
          (native.lineCount as number) < 1 ||
          !Number.isInteger(native.startPage) ||
          !Number.isInteger(native.endPage) ||
          (native.startPage as number) < 1 ||
          (native.startPage as number) > (native.endPage as number) ||
          (native.endPage as number) > (pageCount as number) ||
          typeof native.finalLineText !== "string"
        ) {
          paragraphError = failParagraphs(
            "Word adapter returned invalid paragraph measurements",
          );
          break;
        }
        mapped.push({
          source: "word",
          index: manifest.index,
          position: manifest.position,
          startPage: native.startPage as number,
          endPage: native.endPage as number,
          lineCount: native.lineCount as number,
          finalLineText: native.finalLineText,
          preview: manifest.preview,
        });
      }
    }
    paragraphDiagnostics = paragraphError
      ? { status: "error", error: paragraphError }
      : { status: "ok", value: mapped };
  } else if (paragraphManifest) {
    paragraphDiagnostics = { status: "ok", value: [] };
  }
  return {
    pageCount: pageCount as number,
    totalBodyLines: totalBodyLines as number,
    bodyLinesByPage: bodyLinesByPage as number[],
    bodyLinesOnLastPage: bodyLinesOnLastPage as number | null,
    version: summary.version as string,
    build: summary.build as string,
    activePrinter: summary.activePrinter as string,
    requestedFontFamilies,
    ...(paragraphDiagnostics ? { paragraphDiagnostics } : {}),
    generatedDocxSha256: sha(docx),
    durationMs: performance.now() - start,
    cleanupState: "complete",
  };
}
