import type { CliRuntime } from "./cli-run.js";
import { AgentDocxError } from "./types.js";
export const MAX_JSONL_LINE_BYTES = 8 * 1024 * 1024;

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
    yield* strictUtf8(await runtime.readStdin()).split(/\r?\n/);
    return;
  }

  const decoder = new TextDecoder("utf-8", { fatal: true });
  let pending = "";
  const completeLines = (text: string): string[] => {
    pending += text;
    const lines: string[] = [];
    let newline = pending.indexOf("\n");
    while (newline !== -1) {
      const line = pending.slice(0, newline);
      lines.push(line.endsWith("\r") ? line.slice(0, -1) : line);
      pending = pending.slice(newline + 1);
      newline = pending.indexOf("\n");
    }
    return lines;
  };

  for await (const bytes of runtime.readStdinChunks())
    yield* completeLines(decodeUtf8Chunk(decoder, bytes));
  yield* completeLines(decodeUtf8Chunk(decoder));
  if (pending) yield pending;
}
