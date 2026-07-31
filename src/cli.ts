#!/usr/bin/env node
import { createRequire } from "node:module";
import { runCli, type CliRuntime } from "./cli-run.js";

const { version } = createRequire(import.meta.url)("../package.json") as {
  version: string;
};

async function* readStdinChunks(): AsyncGenerator<Uint8Array> {
  for await (const chunk of process.stdin) {
    yield Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk);
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
