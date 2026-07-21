import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { readFile } from "node:fs/promises";
import test from "node:test";
import { fileURLToPath } from "node:url";
import { parseCliArgs } from "../dist/cli-args.js";
import { runCli } from "../dist/cli-run.js";

const root = fileURLToPath(new URL("..", import.meta.url));
const cli = fileURLToPath(new URL("../dist/cli.js", import.meta.url));
const pkg = JSON.parse(
  await readFile(new URL("../package.json", import.meta.url), "utf8"),
);

function memoryRuntime(input = "", overrides = {}) {
  const stdout = [];
  const stderr = [];
  const runtime = {
    cwd: root,
    stdinIsTTY: false,
    version: pkg.version,
    readStdin: async () =>
      typeof input === "string" ? new TextEncoder().encode(input) : input,
    writeStdout: async (text) => void stdout.push(text),
    writeStderr: async (text) => void stderr.push(text),
    onceSignal() {},
    ...overrides,
  };
  return {
    runtime,
    stdout: () => stdout.join(""),
    stderr: () => stderr.join(""),
  };
}

async function runInProcess(args, input = "", overrides = {}) {
  const capture = memoryRuntime(input, overrides);
  const code = await runCli(args, capture.runtime);
  return { code, stdout: capture.stdout(), stderr: capture.stderr() };
}

function runSubprocess(args, input = "") {
  const { promise, resolve, reject } = Promise.withResolvers();
  const child = spawn(process.execPath, [cli, ...args], {
    cwd: root,
    stdio: ["pipe", "pipe", "pipe"],
  });
  let stdout = "";
  let stderr = "";
  child.stdout.setEncoding("utf8").on("data", (chunk) => (stdout += chunk));
  child.stderr.setEncoding("utf8").on("data", (chunk) => (stderr += chunk));
  child.once("error", reject);
  child.once("close", (code) => resolve({ code, stdout, stderr }));
  child.stdin.end(input);
  return promise;
}

test("help and version are standalone", async () => {
  const help = await runInProcess(["--help"]);
  assert.equal(help.code, 0);
  assert.match(help.stdout, /Usage: md-page-count/);
  const version = await runInProcess(["--version"]);
  assert.equal(version.stdout, `${pkg.version}\n`);
  const bad = await runInProcess(["--help", "x.md"]);
  assert.equal(bad.code, 2);
  assert.match(bad.stderr, /INVALID_ARGUMENT/);
});

test("single JSON is clean and strict UTF-8", async () => {
  const result = await runInProcess(["--json"], "A short filing.\n");
  assert.equal(result.code, 0);
  assert.equal(result.stderr, "");
  assert.equal(JSON.parse(result.stdout).schemaVersion, 1);

  const invalid = await runInProcess(["--json"], new Uint8Array([0xff]));
  assert.equal(invalid.code, 1);
  assert.match(invalid.stderr, /INPUT_NOT_UTF8/);
});

test("boundary fixture through committed config", async () => {
  const result = await runInProcess([
    "test/fixtures/28-hard-lines.md",
    "--config",
    "test/fixtures/exact-27-lines.json",
    "--json",
  ]);
  assert.equal(result.code, 0, result.stderr);
  const measurement = JSON.parse(result.stdout);
  assert.equal(measurement.pageCount, 2);
  assert.equal(measurement.deterministic.equivalentPages, 1 + 24 / 648);
});

test("usage grammar rejects leading-zero counts", async () => {
  const result = await runInProcess(["--page-limit", "01", "--json"], "Text");
  assert.equal(result.code, 2);
  assert.match(result.stderr, /INVALID_ARGUMENT/);
});

test("positional batch preserves order and resets sequence per run", async () => {
  const args = [
    "--batch",
    "test/fixtures/27-hard-lines.md",
    "test/fixtures/28-hard-lines.md",
  ];
  const first = await runInProcess(args);
  assert.equal(first.code, 0, first.stderr);
  const records = first.stdout.trim().split("\n").map(JSON.parse);
  assert.equal(records.length, 2);
  assert.equal(records[0].source.path, "test/fixtures/27-hard-lines.md");
  assert.equal(records[0].sequence, 1);
  assert.equal(records[1].sequence, 2);

  const second = await runInProcess(args);
  assert.equal(JSON.parse(second.stdout.split("\n")[0]).sequence, 1);
});

test("explicit missing LibreOffice does not fall back", async () => {
  const result = await runInProcess(
    [
      "--renderer",
      "libreoffice",
      "--libreoffice-path",
      "/definitely/missing/soffice",
      "--json",
    ],
    "Text",
  );
  assert.equal(result.code, 4);
  assert.match(result.stderr, /LIBREOFFICE_NOT_FOUND/);
});

test("parser rejects duplicate and conflicting mode options", () => {
  assert.throws(
    () => parseCliArgs(["--json", "--json"]),
    /Duplicate option: --json/,
  );
  assert.throws(
    () => parseCliArgs(["--batch", "--json", "file.md"]),
    /Invalid batch option combination/,
  );
});

test("JSONL batch continues after structured item errors", async () => {
  const input = [
    JSON.stringify({ id: "first", markdown: "First." }),
    "not json",
    JSON.stringify({ id: 3, markdown: "Third." }),
    "",
  ].join("\n");
  const result = await runInProcess(["--batch", "--input-jsonl"], input);
  assert.equal(result.code, 1);
  assert.equal(result.stderr, "");
  const records = result.stdout.trim().split("\n").map(JSON.parse);
  assert.deepEqual(
    records.map(({ sequence }) => sequence),
    [1, 2, 3],
  );
  assert.equal(records[0].requestId, "first");
  assert.equal(records[1].kind, "error");
  assert.equal(records[1].error.code, "INVALID_ARGUMENT");
  assert.equal(records[2].requestId, 3);
});

test("executable adapter wires real stdin and clean streams", async () => {
  const result = await runSubprocess(["--json"], "A short filing.\n");
  assert.equal(result.code, 0);
  assert.equal(result.stderr, "");
  assert.equal(JSON.parse(result.stdout).schemaVersion, 1);
});
