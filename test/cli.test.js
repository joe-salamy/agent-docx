import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import {
  mkdir,
  mkdtemp,
  readFile,
  rm,
  symlink,
  writeFile,
} from "node:fs/promises";
import test from "node:test";
import { fileURLToPath } from "node:url";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { parseCliArgs } from "../dist/cli-args.js";
import { runCli, writeOutputExclusive } from "../dist/cli-run.js";
import { inspectDocxTemplate } from "../dist/index.js";

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
  assert.match(help.stdout, /Usage: agent-docx/);
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

test("single output writes valid exclusive DOCX without JSON bytes", async () => {
  const temporary = await mkdtemp(join(tmpdir(), "agent-docx-output-"));
  const output = join(temporary, "result");
  try {
    const result = await runInProcess(
      ["--json", "--output", output],
      "# Output\n",
    );
    assert.equal(result.code, 0, result.stderr);
    assert.equal(result.stderr, "");
    assert.equal("generatedDocx" in JSON.parse(result.stdout), false);
    const bytes = await readFile(output);
    assert.equal(bytes.subarray(0, 2).toString(), "PK");
    await inspectDocxTemplate(bytes);

    const original = Buffer.from(bytes);
    const collision = await runInProcess(
      ["--json", "--output", output],
      "# Changed\n",
    );
    assert.equal(collision.code, 1);
    assert.equal(collision.stdout, "");
    assert.deepEqual(await readFile(output), original);
    assert.deepEqual(JSON.parse(collision.stderr), {
      error: {
        code: "OUTPUT_EXISTS",
        message: `Output already exists: ${output}`,
      },
    });

    const missing = join(temporary, "missing", "result.docx");
    const failed = await runInProcess(
      ["--json", "--output", missing],
      "# Missing parent\n",
    );
    assert.equal(failed.code, 1);
    assert.equal(failed.stdout, "");
    const error = JSON.parse(failed.stderr).error;
    assert.equal(error.code, "OUTPUT_WRITE_FAILED");
    assert.equal(error.message, `Failed to write output: ${missing}`);
    assert.equal(typeof error.details.cause, "string");
  } finally {
    await rm(temporary, { recursive: true, force: true });
  }
});

test("exclusive output cleanup preserves the original write or close error", async () => {
  for (const phase of ["write", "close"]) {
    const calls = [];
    let closeCalls = 0;
    await assert.rejects(
      () =>
        writeOutputExclusive("/resolved", "shown.docx", new Uint8Array([1]), {
          async open(path, flags) {
            calls.push(["open", path, flags]);
            return {
              async writeFile() {
                calls.push(["write"]);
                if (phase === "write") throw new Error("write-original");
              },
              async close() {
                closeCalls++;
                calls.push(["close"]);
                if (phase === "close" || phase === "write") {
                  throw new Error(
                    phase === "close" ? "close-original" : "cleanup-close",
                  );
                }
              },
            };
          },
          async unlink(path) {
            calls.push(["unlink", path]);
            throw new Error("cleanup-unlink");
          },
        }),
      (error) => {
        assert.equal(error.code, "OUTPUT_WRITE_FAILED");
        assert.equal(
          error.details.cause,
          phase === "write" ? "write-original" : "close-original",
        );
        return true;
      },
    );
    assert.ok(closeCalls >= 1);
    assert.ok(calls.some(([name]) => name === "unlink"));
  }
});

test("output and discovery options enforce mode boundaries", () => {
  for (const args of [
    ["--output", "-"],
    ["--help", "--output", "x"],
    ["--inspect-template", "--output", "x", "template.docx"],
    ["--batch", "--output", "x", "a.md"],
    ["--batch", "--input-jsonl", "--include", "*.md"],
    ["--watch", "--output", "x", "a.md"],
    ["--watch", "--recursive", "a.md"],
  ]) {
    assert.throws(
      () => parseCliArgs(args),
      (error) => error.code === "INVALID_ARGUMENT",
    );
  }
  assert.throws(
    () => parseCliArgs(["--batch", "--recursive", "--no-recursive", "x"]),
    /cannot be combined/,
  );
  assert.throws(
    () => parseCliArgs(["--batch", "-"]),
    /Positional batch does not accept stdin/,
  );
  assert.doesNotThrow(() =>
    parseCliArgs([
      "--batch",
      "--include",
      "*.md",
      "--include",
      "*.txt",
      "--exclude",
      "skip.md",
      "docs",
    ]),
  );
});

test("sections appear only when requested in human, JSON, and batch output", async () => {
  const human = await runInProcess(
    ["--sections", "--page-limit", "1"],
    "# Analysis\n\nBody.\n\n<!-- pagebreak -->\n\nMore.",
  );
  assert.equal(human.code, 0, human.stderr);
  assert.match(
    human.stdout,
    /Section 0 \(preamble\): 0 pages; 0 visual lines; 0 counted lines/,
  );
  assert.match(human.stdout, /Section 1 \(H1 "Analysis"\): pages 1,2 \(2\)/);
  assert.match(human.stdout, /beyond limit 1: 2/);

  const json = await runInProcess(["--sections", "--json"], "# JSON\n\nBody.");
  assert.equal(
    JSON.parse(json.stdout).deterministic.sections[1].source,
    "deterministic",
  );
  const regular = await runInProcess(["--json"], "# JSON\n\nBody.");
  assert.equal("sections" in JSON.parse(regular.stdout).deterministic, false);

  const batch = await runInProcess(
    ["--batch", "--input-jsonl", "--sections"],
    `${JSON.stringify({ markdown: "# Batch\n\nBody." })}\n`,
  );
  assert.equal(
    JSON.parse(batch.stdout).measurement.deterministic.sections[1].source,
    "deterministic",
  );
});

test("positional batch discovers, filters, sorts, and deduplicates snapshots", async () => {
  const temporary = await mkdtemp(join(tmpdir(), "agent-docx-batch-"));
  const docs = join(temporary, "docs");
  await mkdir(join(docs, "nested"), { recursive: true });
  await Promise.all([
    writeFile(join(docs, "a.md"), "A."),
    writeFile(join(docs, "Z.md"), "Z."),
    writeFile(join(docs, "nested", "b.md"), "B."),
    writeFile(join(docs, "nested", "skip.md"), "Skip."),
    writeFile(join(temporary, "explicit.txt"), "Explicit."),
  ]);
  try {
    const result = await runInProcess(
      [
        "--batch",
        "--recursive",
        "--include",
        "*.md",
        "--exclude",
        "skip.md",
        "explicit.txt",
        "docs",
        "docs/**/*.md",
      ],
      "",
      { cwd: temporary },
    );
    assert.equal(result.code, 0, result.stderr);
    const rows = result.stdout.trim().split("\n").map(JSON.parse);
    assert.deepEqual(
      rows.map(({ source }) => source.path),
      ["explicit.txt", "docs/Z.md", "docs/a.md", "docs/nested/b.md"],
    );
    assert.deepEqual(
      rows.map(({ sequence }) => sequence),
      [1, 2, 3, 4],
    );

    const shallow = await runInProcess(
      ["--batch", "--no-recursive", "docs"],
      "",
      { cwd: temporary },
    );
    assert.deepEqual(
      shallow.stdout
        .trim()
        .split("\n")
        .map(JSON.parse)
        .map(({ source }) => source.path),
      ["docs/Z.md", "docs/a.md"],
    );

    const custom = await runInProcess(
      ["--batch", "--include", "*.txt", "explicit.txt"],
      "",
      { cwd: temporary },
    );
    assert.equal(custom.code, 0);
    assert.equal(JSON.parse(custom.stdout).source.path, "explicit.txt");

    try {
      await symlink(join(docs, "a.md"), join(temporary, "alias.md"));
      const aliases = await runInProcess(
        ["--batch", "alias.md", "docs/a.md"],
        "",
        { cwd: temporary },
      );
      assert.equal(aliases.stdout.trim().split("\n").length, 1);
    } catch (error) {
      if (!["EPERM", "EACCES", "ENOTSUP"].includes(error.code)) throw error;
    }
  } finally {
    await rm(temporary, { recursive: true, force: true });
  }
});

test("batch discovery preflights failures and config overrides", async () => {
  const temporary = await mkdtemp(
    join(tmpdir(), "agent-docx-batch-config-"),
  );
  await mkdir(join(temporary, "docs"));
  await writeFile(join(temporary, "docs", "a.md"), "A.");
  await writeFile(join(temporary, "docs", "a.txt"), "Text.");
  try {
    for (const selector of ["docs/no-match-*.md", "docs/[broken"]) {
      const result = await runInProcess(["--batch", selector], "", {
        cwd: temporary,
      });
      assert.equal(result.code, 2);
      assert.equal(result.stdout, "");
      assert.match(result.stderr, /INVALID_ARGUMENT/);
    }
    const filtered = await runInProcess(
      ["--batch", "--exclude", "*.md", "docs"],
      "",
      { cwd: temporary },
    );
    assert.equal(filtered.code, 2);
    assert.equal(filtered.stdout, "");

    const missing = await runInProcess(
      ["--batch", "missing.md", "docs/a.md"],
      "",
      { cwd: temporary },
    );
    assert.equal(missing.code, 1);
    const missingRows = missing.stdout.trim().split("\n").map(JSON.parse);
    assert.equal(missingRows[0].kind, "error");
    assert.equal(missingRows[1].kind, "result");

    const configPath = join(temporary, "config.json");
    await writeFile(
      configPath,
      JSON.stringify({
        sectionDiagnostics: true,
        batch: { recursive: false, include: ["*.txt"] },
      }),
    );
    const configured = await runInProcess(
      ["--batch", "--config", configPath, "docs"],
      "",
      { cwd: temporary },
    );
    assert.equal(JSON.parse(configured.stdout).source.path, "docs/a.txt");
    const replaced = await runInProcess(
      ["--batch", "--config", configPath, "--include", "*.md", "docs"],
      "",
      { cwd: temporary },
    );
    assert.equal(JSON.parse(replaced.stdout).source.path, "docs/a.md");

    await writeFile(configPath, JSON.stringify({ batch: { include: [] } }));
    const invalidConfig = await runInProcess(
      ["--batch", "--config", configPath, "docs"],
      "",
      { cwd: temporary },
    );
    assert.equal(invalidConfig.code, 1);
    assert.match(invalidConfig.stderr, /INVALID_CONFIG/);
  } finally {
    await rm(temporary, { recursive: true, force: true });
  }
});
