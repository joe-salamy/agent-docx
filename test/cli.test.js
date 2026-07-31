import assert from "node:assert/strict";
import { Ajv2020 } from "ajv/dist/2020.js";
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
const schemaNames = [
  "cli-request.schema.json",
  "measurement-result.schema.json",
  "docx-template-inspection.schema.json",
  "cli-jsonl.schema.json",
  "cli-error.schema.json",
  "profile-catalog.schema.json",
];
const schemas = await Promise.all(
  schemaNames.map(async (name) =>
    JSON.parse(await readFile(new URL(`../${name}`, import.meta.url), "utf8")),
  ),
);
const ajv = new Ajv2020({ strict: true, allowUnionTypes: true });
for (const schema of schemas) ajv.addSchema(schema);
const validateMeasurement = ajv.getSchema(
  "https://md-page-count.dev/schemas/measurement-result-v1.json",
);
const validateRequest = ajv.getSchema(
  "https://md-page-count.dev/schemas/cli-request-v1.json",
);
const validateJsonl = ajv.getSchema(
  "https://md-page-count.dev/schemas/cli-jsonl-v1.json",
);
const validateInspection = ajv.getSchema(
  "https://md-page-count.dev/schemas/docx-template-inspection-v1.json",
);
const validateFatal = ajv.getSchema(
  "https://md-page-count.dev/schemas/cli-error-v1.json",
);
const validateProfileCatalog = ajv.getSchema(
  "https://md-page-count.dev/schemas/profile-catalog-v1.json",
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

test("profile catalog is standalone and schema-valid", async () => {
  const json = await runInProcess(["--list-profiles", "--json"]);
  assert.equal(json.code, 0, json.stderr);
  assert.equal(json.stderr, "");
  const catalog = JSON.parse(json.stdout);
  assert.deepEqual(
    catalog.profiles.map(({ id }) => id),
    ["us-district-conventional", "frap-32", "cand-civil"],
  );
  assert.equal(
    validateProfileCatalog(catalog),
    true,
    JSON.stringify(validateProfileCatalog.errors),
  );

  const human = await runInProcess(["--list-profiles"]);
  assert.equal(human.code, 0, human.stderr);
  assert.match(human.stdout, /Built-in profiles:/);
  assert.match(human.stdout, /frap-32: Federal Rule/);

  const invalid = await runInProcess(["--list-profiles", "filing.md"]);
  assert.equal(invalid.code, 2);
  assert.match(invalid.stderr, /INVALID_ARGUMENT/);
});

test("single JSON is clean and strict UTF-8", async () => {
  const result = await runInProcess(["--json"], "A short filing.\n");
  assert.equal(result.code, 0);
  assert.equal(result.stderr, "");
  assert.equal(JSON.parse(result.stdout).schemaVersion, 1);
  assert.equal(
    validateMeasurement(JSON.parse(result.stdout)),
    true,
    JSON.stringify(validateMeasurement.errors),
  );

  const invalid = await runInProcess(["--json"], new Uint8Array([0xff]));
  assert.equal(invalid.code, 1);
  assert.match(invalid.stderr, /INPUT_NOT_UTF8/);
});

test("template inspection JSON satisfies its published schema", async () => {
  const result = await runInProcess([
    "--inspect-template",
    "test/fixtures/docx/theme-inheritance.docx",
    "--json",
  ]);
  assert.equal(result.code, 0, result.stderr);
  assert.equal(result.stderr, "");
  assert.equal(
    validateInspection(JSON.parse(result.stdout)),
    true,
    JSON.stringify(validateInspection.errors),
  );
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
  for (const record of records)
    assert.equal(
      validateJsonl(record),
      true,
      JSON.stringify(validateJsonl.errors),
    );
});

test("JSONL batch decodes streamed requests across UTF-8 boundaries", async () => {
  const input = [
    JSON.stringify({
      id: "first",
      name: "résumé",
      markdown: "First 😀.",
    }),
    JSON.stringify({ id: "second", markdown: "Second." }),
    "",
  ].join("\r\n");
  const bytes = Buffer.from(input);
  const emojiStart = bytes.indexOf(Buffer.from("😀"));
  const newline = bytes.indexOf(0x0a);
  assert.ok(emojiStart >= 0);
  assert.ok(newline > emojiStart);
  const chunks = [
    bytes.subarray(0, emojiStart + 2),
    bytes.subarray(emojiStart + 2, newline + 1),
    bytes.subarray(newline + 1),
  ];
  const result = await runInProcess(["--batch", "--input-jsonl"], "", {
    readStdin: async () => {
      throw new Error("JSONL batch must use streaming stdin");
    },
    readStdinChunks: async function* () {
      for (const chunk of chunks) yield chunk;
    },
  });
  assert.equal(result.code, 0, result.stderr);
  const records = result.stdout.trim().split("\n").map(JSON.parse);
  assert.deepEqual(
    records.map(({ requestId }) => requestId),
    ["first", "second"],
  );
  assert.deepEqual(records[0].source, { kind: "inline", name: "résumé" });
  for (const record of records)
    assert.equal(
      validateJsonl(record),
      true,
      JSON.stringify(validateJsonl.errors),
    );
});

test("request schema accepts exactly one closed source shape", () => {
  assert.equal(validateRequest({ id: "x", markdown: "Text." }), true);
  assert.equal(validateRequest({ id: 1, path: "filing.md" }), true);
  assert.equal(validateRequest({ markdown: "Text.", extra: true }), false);
  assert.equal(
    validateRequest({ markdown: "Text.", path: "filing.md" }),
    false,
  );
  assert.equal(validateRequest({ id: {}, markdown: "Text." }), false);
});

test("JSONL rejects closed requests while retaining valid correlation and relative sources", async () => {
  const input = [
    JSON.stringify({
      id: "known",
      markdown: "Inline.",
      unexpected: true,
    }),
    JSON.stringify({ id: 7, path: "../outside.md", extra: 1 }),
    JSON.stringify({ id: {}, markdown: "Bad ID." }),
    JSON.stringify({ id: "empty", path: "" }),
  ].join("\n");
  const result = await runInProcess(["--batch", "--input-jsonl"], input);
  assert.equal(result.code, 1);
  assert.equal(result.stderr, "");
  const records = result.stdout.trim().split("\n").map(JSON.parse);
  assert.deepEqual(
    records.map(({ requestId }) => requestId),
    ["known", 7, null, "empty"],
  );
  assert.deepEqual(records[0].source, { kind: "inline", name: null });
  assert.deepEqual(records[1].source, {
    kind: "file",
    path: "../outside.md",
  });
  assert.deepEqual(records[2].source, { kind: "inline", name: null });
  assert.deepEqual(records[3].source, { kind: "stdin" });
  for (const record of records)
    assert.equal(
      validateJsonl(record),
      true,
      JSON.stringify(validateJsonl.errors),
    );
});

test("trim-only human output is ranked and omits generic paragraph bars", async () => {
  const trim = await runInProcess(
    ["--trim", "--trim-threshold", "1"],
    "Word ".repeat(250),
  );
  assert.equal(trim.code, 0, trim.stderr);
  assert.match(trim.stdout, /Trim opportunities:/);
  assert.match(trim.stdout, /1\. Lines \d+-\d+: "/);
  assert.match(trim.stdout, /\d+ twips/);
  assert.doesNotMatch(trim.stdout, /[█░]/);

  const paragraphs = await runInProcess(
    ["--trim", "--trim-threshold", "1", "--paragraphs"],
    "Word ".repeat(250),
  );
  assert.match(paragraphs.stdout, /[█░]/);
});

test("fatal records use INTERNAL_ERROR for unknown exceptions", async () => {
  const capture = memoryRuntime("", {
    readStdin: async () => {
      throw new Error("injected failure");
    },
  });
  const code = await runCli(["--json"], capture.runtime);
  assert.equal(code, 1);
  const fatal = JSON.parse(capture.stderr());
  assert.equal(
    validateFatal(fatal),
    true,
    JSON.stringify(validateFatal.errors),
  );
  assert.equal(fatal.error.code, "INTERNAL_ERROR");
  assert.equal(fatal.error.message, "injected failure");
});

test("executable adapter wires real stdin and clean streams", async () => {
  const result = await runSubprocess(["--json"], "A short filing.\n");
  assert.equal(result.code, 0);
  assert.equal(result.stderr, "");
  assert.equal(JSON.parse(result.stdout).schemaVersion, 1);
});

test("single output writes valid exclusive DOCX without JSON bytes", async () => {
  const temporary = await mkdtemp(join(tmpdir(), "md-page-count-output-"));
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
      schemaVersion: 1,
      kind: "fatal",
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
  const temporary = await mkdtemp(join(tmpdir(), "md-page-count-batch-"));
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
    join(tmpdir(), "md-page-count-batch-config-"),
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

test("watch JSONL emits validated ready, result, and signal end records", async () => {
  const temporary = await mkdtemp(join(tmpdir(), "md-page-count-watch-"));
  const path = join(temporary, "watch.md");
  await writeFile(path, "Watched.");
  try {
    const capture = memoryRuntime("", {
      cwd: temporary,
      onceSignal(signal, listener) {
        if (signal === "SIGINT") setTimeout(listener, 10);
      },
    });
    const code = await runCli(
      ["--watch", "--jsonl", "watch.md"],
      capture.runtime,
    );
    assert.equal(code, 130);
    assert.equal(capture.stderr(), "");
    const records = capture.stdout().trim().split("\n").map(JSON.parse);
    assert.deepEqual(
      records.map(({ kind }) => kind),
      ["ready", "result", "end"],
    );
    assert.deepEqual(
      records.map(({ sequence }) => sequence),
      [1, 2, 3],
    );
    assert.equal(records[1].trigger.kind, "initial");
    assert.equal(records[2].reason, "SIGINT");
    for (const record of records)
      assert.equal(
        validateJsonl(record),
        true,
        JSON.stringify(validateJsonl.errors),
      );
  } finally {
    await rm(temporary, { recursive: true, force: true });
  }
});
