import type { CliRuntime } from "./cli-run.js";
import { AgentDocxError } from "./types.js";
export const MAX_JSONL_LINE_BYTES = 8 * 1024 * 1024;

const jsonlLineTooLarge = (): AgentDocxError =>
  new AgentDocxError(
    "INVALID_ARGUMENT",
    `JSONL line exceeds ${MAX_JSONL_LINE_BYTES} bytes`,
  );

const checkJsonlLineSize = (line: string): string => {
  if (Buffer.byteLength(line) > MAX_JSONL_LINE_BYTES)
    throw jsonlLineTooLarge();
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
    for (const line of strictUtf8(await runtime.readStdin()).split(/\r?\n/))
      yield checkJsonlLineSize(line);
    return;
  }

  const decoder = new TextDecoder("utf-8", { fatal: true });
  let pending = "";
  const completeLines = (text: string): string[] => {
    pending += text;
    if (Buffer.byteLength(pending) > MAX_JSONL_LINE_BYTES)
      throw jsonlLineTooLarge();
    const lines: string[] = [];
    let newline = pending.indexOf("\n");
    while (newline !== -1) {
      const line = pending.slice(0, newline);
      lines.push(
        checkJsonlLineSize(line.endsWith("\r") ? line.slice(0, -1) : line),
      );
      pending = pending.slice(newline + 1);
      newline = pending.indexOf("\n");
    }
    return lines;
  };

  for await (const bytes of runtime.readStdinChunks())
    yield* completeLines(decodeUtf8Chunk(decoder, bytes));
  yield* completeLines(decodeUtf8Chunk(decoder));
  if (pending) yield checkJsonlLineSize(pending);
}
