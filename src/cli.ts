#!/usr/bin/env node
import { readFileSync } from "node:fs";
import { MAX_INPUT_BYTES } from "./input.js";
import { AgentDocxError } from "./types.js";
import { runCli, type CliRuntime } from "./cli-run.js";

const version: string = (() => {
  try {
    const text = readFileSync(new URL("../package.json", import.meta.url), "utf8");
    return (JSON.parse(text) as { version: string }).version;
  } catch {
    return "0.1.1";
  }
})();

async function* readStdinChunks(): AsyncGenerator<Uint8Array> {
  let total = 0;
  for await (const chunk of process.stdin) {
    const bytes = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk);
    total += bytes.byteLength;
    if (total > MAX_INPUT_BYTES)
      throw new AgentDocxError(
        "INPUT_TOO_LARGE",
        `stdin exceeds the ${MAX_INPUT_BYTES} byte input limit`,
      );
    yield bytes;
  }
}

async function readStdin(): Promise<Uint8Array> {
  const chunks: Uint8Array[] = [];
  for await (const chunk of readStdinChunks()) chunks.push(chunk);
  return Buffer.concat(chunks);
}

async function writeStream(
  stream: NodeJS.WriteStream,
  text: string,
): Promise<void> {
  if (!stream.write(text)) {
    await new Promise<void>((resolve) => stream.once("drain", resolve));
  }
}

const runtime: CliRuntime = {
  cwd: process.cwd(),
  stdinIsTTY: process.stdin.isTTY === true,
  version,
  readStdin,
  readStdinChunks,
  writeStdout: (text) => writeStream(process.stdout, text),
  writeStderr: (text) => writeStream(process.stderr, text),
  onceSignal: (signal, listener) => process.once(signal, listener),
};

process.stdout.on("error", (error) => {
  if ((error as NodeJS.ErrnoException).code === "EPIPE") process.exitCode = 0;
});

runCli(process.argv.slice(2), runtime)
  .then((code) => {
    process.exitCode = code;
  })
  .catch(() => {
    process.exitCode = 1;
  });
