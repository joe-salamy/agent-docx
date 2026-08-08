import type { CliRuntime } from "./cli-run.js";
import { MAX_INPUT_BYTES } from "./input.js";
import { AgentDocxError } from "./types.js";
export const MAX_JSONL_LINE_BYTES = 8 * 1024 * 1024;

const jsonlLineTooLarge = (): AgentDocxError =>
  new AgentDocxError(
    "INVALID_ARGUMENT",
    `JSONL line exceeds ${MAX_JSONL_LINE_BYTES} bytes`,
  );

const inputTooLarge = (): AgentDocxError =>
  new AgentDocxError(
    "INPUT_TOO_LARGE",
    `stdin exceeds the ${MAX_INPUT_BYTES} byte input limit`,
  );

const checkJsonlLineSize = (line: string): string => {
  if (Buffer.byteLength(line) > MAX_JSONL_LINE_BYTES) throw jsonlLineTooLarge();
  return line;
};

export const strictUtf8 = (bytes: Uint8Array): string => {
  try {
    return new TextDecoder("utf-8", { fatal: true }).decode(bytes);
  } catch {
    throw new AgentDocxError("INPUT_NOT_UTF8", "Input is not valid UTF-8");
  }
};

const decodeUtf8Chunk = (decoder: TextDecoder, bytes?: Uint8Array): string => {
  try {
    return bytes ? decoder.decode(bytes, { stream: true }) : decoder.decode();
  } catch {
    throw new AgentDocxError("INPUT_NOT_UTF8", "Input is not valid UTF-8");
  }
};

/** Decodes JSONL input incrementally while preserving UTF-8 code points across chunks. */
export async function* jsonlLines(
  runtime: Pick<CliRuntime, "readStdin" | "readStdinChunks">,
): AsyncGenerator<string, void, undefined> {
  if (!runtime.readStdinChunks) {
    const bytes = await runtime.readStdin();
    if (bytes.byteLength > MAX_INPUT_BYTES) throw inputTooLarge();
    for (const line of strictUtf8(bytes).split(/\r?\n/))
      yield checkJsonlLineSize(line);
    return;
  }

  const decoder = new TextDecoder("utf-8", { fatal: true });
  let pending = "";
  let total = 0;
  const completeLines = (text: string): string[] => {
    const lines: string[] = [];
    let start = 0;
    while (start <= text.length) {
      const newline = text.indexOf("\n", start);
      const fragment = text.slice(
        start,
        newline === -1 ? text.length : newline,
      );
      const rawBytes = Buffer.byteLength(pending) + Buffer.byteLength(fragment);
      const normalizedBytes =
        newline !== -1 && fragment.endsWith("\r") ? rawBytes - 1 : rawBytes;
      if (normalizedBytes > MAX_JSONL_LINE_BYTES) throw jsonlLineTooLarge();
      if (newline === -1) {
        pending += fragment;
        break;
      }
      const rawLine = pending + fragment;
      pending = "";
      lines.push(
        checkJsonlLineSize(
          rawLine.endsWith("\r") ? rawLine.slice(0, -1) : rawLine,
        ),
      );
      start = newline + 1;
      if (start === text.length) break;
    }
    return lines;
  };

  for await (const bytes of runtime.readStdinChunks()) {
    total += bytes.byteLength;
    if (total > MAX_INPUT_BYTES) throw inputTooLarge();
    yield* completeLines(decodeUtf8Chunk(decoder, bytes));
  }
  yield* completeLines(decodeUtf8Chunk(decoder));
  if (pending) yield checkJsonlLineSize(pending);
}
