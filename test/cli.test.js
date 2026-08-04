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
import { parseAgentRequest, serializeAgentValue } from "../dist/agent.js";

const root = fileURLToPath(new URL("..", import.meta.url));
const cli = fileURLToPath(new URL("../dist/cli.js", import.meta.url));
const pkg = JSON.parse(
  await readFile(new URL("../package.json", import.meta.url), "utf8"),
);
const schemaNames = [
  "measurement-request.schema.json",
  "measurement-result.schema.json",
  "docx-template-inspection.schema.json",
  "measurement-stream.schema.json",
  "cli-error.schema.json",
  "agent-stream.schema.json",
  "agent-request.schema.json",
  "agent-response.schema.json",
  "project.schema.json",
  "rule-pack.schema.json",
  "revision.schema.json",
  "change-set.schema.json",
  "source-patch.schema.json",
  "validation-result.schema.json",
  "artifact-result.schema.json",
  "compiled-docx.schema.json",
  "docx-import-result.schema.json",
  "redline-import-result.schema.json",
  "filing-set.schema.json",
  "filing-set-validation.schema.json",
  "profile-catalog.schema.json",
];
const schemas = await Promise.all(
  schemaNames.map(async (name) =>
    JSON.parse(await readFile(new URL(`../${name}`, import.meta.url), "utf8")),
  ),
);
const ajv = new Ajv2020({
  strict: true,
  allowUnionTypes: true,
  formats: { date: true, "date-time": true, uri: true },
});
for (const schema of schemas) ajv.addSchema(schema);
const validateMeasurement = ajv.getSchema(
  "https://agent-docx.dev/schemas/measurement-result-v1.json",
);
const validateRequest = ajv.getSchema(
  "https://agent-docx.dev/schemas/measurement-request-v1.json",
);
const validateJsonl = ajv.getSchema(
  "https://agent-docx.dev/schemas/measurement-stream-v1.json",
);
const validateInspection = ajv.getSchema(
  "https://agent-docx.dev/schemas/docx-template-inspection-v1.json",
);
const validateFatal = ajv.getSchema(
  "https://agent-docx.dev/schemas/cli-error-v1.json",
);
const validateProfileCatalog = ajv.getSchema(
  "https://agent-docx.dev/schemas/profile-catalog-v1.json",
);
const validateAgentStream = ajv.getSchema(
  "https://agent-docx.dev/schemas/agent-stream-v1.json",
);
const validateAgentRequest = ajv.getSchema(
  "https://agent-docx.dev/schemas/agent-request-v1.json",
);
const validateAgentResponse = ajv.getSchema(
  "https://agent-docx.dev/schemas/agent-response-v1.json",
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

test("explicit project and agent commands execute a revision-bound workflow", async () => {
  const directory = await mkdtemp(join(tmpdir(), "agent-docx-cli-"));
  const manifest = join(directory, "agent-docx.json");
  try {
    await writeFile(join(directory, "motion.md"), "# Motion\n\nBody text.\n");
    await writeFile(
      join(directory, "metadata.json"),
      JSON.stringify({
        court: "United States District Court",
        jurisdiction: "Northern District of California",
        caseName: "Example v. Example",
        docketNumber: "3:26-cv-00001",
        documentTitle: "Motion",
        parties: [],
        counsel: [],
        certificates: [],
      }),
    );
    const initialized = await runInProcess(
      [
        "project",
        "init",
        "--project",
        "agent-docx.json",
        "--document",
        "motion",
        "--source",
        "motion.md",
        "--profile",
        "us-district-conventional",
        "--metadata",
        "metadata.json",
        "--json",
      ],
      "",
      { cwd: directory },
    );
    assert.equal(initialized.code, 0, initialized.stderr);
    assert.equal(
      JSON.parse(initialized.stdout).manifest.defaultDocument,
      "motion",
    );
    const checkpoint = await runInProcess(
      [
        "revision",
        "checkpoint",
        "--project",
        "agent-docx.json",
        "--document",
        "motion",
        "--author",
        "Drafter",
        "--message",
        "Initial draft",
        "--json",
      ],
      "",
      { cwd: directory },
    );
    assert.equal(checkpoint.code, 0, checkpoint.stderr);
    const revision = JSON.parse(checkpoint.stdout).revision.id;
    const agent = await runInProcess(
      ["agent", "--input-jsonl"],
      `${JSON.stringify({
        schemaVersion: 1,
        id: "state",
        action: "document.get",
        project: manifest,
        params: { documentId: "motion", revision: "HEAD" },
      })}\n`,
      { cwd: directory },
    );
    assert.equal(agent.code, 0, agent.stderr);
    const response = JSON.parse(agent.stdout);
    assert.equal(response.kind, "result");
    assert.equal(response.requestId, "state");
    assert.equal(response.revision, revision);
    assert.match(response.value.source, /agent-docx:block/);
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});

test("agent JSONL decodes streamed requests across UTF-8 boundaries", async () => {
  const input = `${JSON.stringify({
    schemaVersion: 1,
    id: "emoji-😀",
    action: "project.get",
    params: {},
  })}\n`;
  const bytes = Buffer.from(input);
  const emoji = bytes.indexOf(Buffer.from("😀"));
  assert.ok(emoji >= 0);
  const result = await runInProcess(["agent", "--input-jsonl"], "", {
    readStdin: async () => {
      throw new Error("agent JSONL must use streaming stdin");
    },
    readStdinChunks: async function* () {
      yield bytes.subarray(0, emoji + 2);
      yield bytes.subarray(emoji + 2);
    },
  });
  assert.equal(result.code, 0, result.stderr);
  const response = JSON.parse(result.stdout);
  assert.equal(response.kind, "error");
  assert.equal(response.requestId, "emoji-😀");
  assert.equal(
    validateAgentResponse(response),
    true,
    JSON.stringify(validateAgentResponse.errors),
  );
});

test("agent JSONL rejects oversized request lines", async () => {
  const result = await runInProcess(
    ["agent", "--input-jsonl"],
    `x${" ".repeat(8 * 1024 * 1024 + 1)}\n`,
  );
  assert.equal(result.code, 0, result.stderr);
  const response = JSON.parse(result.stdout);
  assert.equal(response.kind, "error");
  assert.equal(response.error.code, "INVALID_ARGUMENT");
  assert.match(response.error.message, /JSONL input line exceeds/);
  assert.equal(validateAgentResponse(response), true);
});

test("help and version are standalone", async () => {
  const help = await runInProcess(["--help"]);
  assert.equal(help.code, 0);
  assert.match(help.stdout, /Usage:\n  agent-docx --help/);
  const version = await runInProcess(["--version"]);
  assert.equal(version.stdout, `${pkg.version}\n`);
  const bad = await runInProcess(["--help", "x.md"]);
  assert.equal(bad.code, 2);
  assert.match(bad.stderr, /INVALID_ARGUMENT/);
});

test("profile catalog is standalone and schema-valid", async () => {
  const json = await runInProcess(["profiles", "--json"]);
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

  const human = await runInProcess(["profiles"]);
  assert.equal(human.code, 0, human.stderr);
  assert.match(human.stdout, /Built-in profiles:/);
  assert.match(human.stdout, /frap-32: Federal Rule/);

  const invalid = await runInProcess(["profiles", "filing.md"]);
  assert.equal(invalid.code, 2);
  assert.match(invalid.stderr, /INVALID_ARGUMENT/);
});

test("single JSON is clean and strict UTF-8", async () => {
  const result = await runInProcess(["measure", "--json"], "A short filing.\n");
  assert.equal(result.code, 0);
  assert.equal(result.stderr, "");
  assert.equal(JSON.parse(result.stdout).schemaVersion, 1);
  assert.equal(
    validateMeasurement(JSON.parse(result.stdout)),
    true,
    JSON.stringify(validateMeasurement.errors),
  );

  const invalid = await runInProcess(
    ["measure", "--json"],
    new Uint8Array([0xff]),
  );
  assert.equal(invalid.code, 1);
  assert.match(invalid.stderr, /INPUT_NOT_UTF8/);
});

test("template inspection JSON satisfies its published schema", async () => {
  const result = await runInProcess([
    "template",
    "inspect",
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
    "measure",
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
  const result = await runInProcess(
    ["measure", "--page-limit", "01", "--json"],
    "Text",
  );
  assert.equal(result.code, 2);
  assert.match(result.stderr, /INVALID_ARGUMENT/);
});

test("positional batch preserves order and resets sequence per run", async () => {
  const args = [
    "measure",
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
      "measure",
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
    () => parseCliArgs(["measure", "--json", "--json"]),
    /Duplicate option: --json/,
  );
  assert.throws(
    () => parseCliArgs(["measure", "--batch", "--json", "file.md"]),
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
  const result = await runInProcess(
    ["measure", "--batch", "--input-jsonl"],
    input,
  );
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
  const result = await runInProcess(
    ["measure", "--batch", "--input-jsonl"],
    "",
    {
      readStdin: async () => {
        throw new Error("JSONL batch must use streaming stdin");
      },
      readStdinChunks: async function* () {
        for (const chunk of chunks) yield chunk;
      },
    },
  );
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

test("agent request schema accepts every supported filing kind", () => {
  const metadata = {
    court: "United States District Court",
    jurisdiction: "Northern District of California",
    caseName: "Example v. Example",
    docketNumber: "3:26-cv-00001",
    documentTitle: "Motion",
    parties: [],
    counsel: [],
    certificates: [],
  };
  for (const filingKind of [
    "principal-brief",
    "reply-brief",
    "motion-document",
    "opposition-text",
    "reply-text",
  ]) {
    const request = {
      schemaVersion: 1,
      id: filingKind,
      action: "project.add",
      project: "agent-docx.json",
      params: {
        documentId: "motion",
        source: "motion.md",
        profile: "us-district-conventional",
        filingKind,
        metadata,
      },
    };
    assert.equal(
      validateAgentRequest(request),
      true,
      JSON.stringify(validateAgentRequest.errors),
    );
  }
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
  const result = await runInProcess(
    ["measure", "--batch", "--input-jsonl"],
    input,
  );
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
    ["measure", "--trim", "--trim-threshold", "1"],
    "Word ".repeat(250),
  );
  assert.equal(trim.code, 0, trim.stderr);
  assert.match(trim.stdout, /Trim opportunities:/);
  assert.match(trim.stdout, /1\. Lines \d+-\d+: "/);
  assert.match(trim.stdout, /\d+ twips/);
  assert.doesNotMatch(trim.stdout, /[█░]/);

  const paragraphs = await runInProcess(
    ["measure", "--trim", "--trim-threshold", "1", "--paragraphs"],
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
  const code = await runCli(["measure", "--json"], capture.runtime);
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
  const result = await runSubprocess(
    ["measure", "--json"],
    "A short filing.\n",
  );
  assert.equal(result.code, 0);
  assert.equal(result.stderr, "");
  assert.equal(JSON.parse(result.stdout).schemaVersion, 1);
});

test("single output writes valid exclusive DOCX without JSON bytes", async () => {
  const temporary = await mkdtemp(join(tmpdir(), "agent-docx-output-"));
  const output = join(temporary, "result");
  try {
    const result = await runInProcess(
      ["measure", "--json", "--output", output],
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
      ["measure", "--json", "--output", output],
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
      ["measure", "--json", "--output", missing],
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
    ["measure", "--output", "-"],
    ["--help", "--output", "x"],
    ["template", "inspect", "--output", "x", "template.docx"],
    ["measure", "--batch", "--output", "x", "a.md"],
    ["measure", "--batch", "--input-jsonl", "--include", "*.md"],
    ["measure", "--watch", "--output", "x", "a.md"],
    ["measure", "--watch", "--recursive", "a.md"],
  ]) {
    assert.throws(
      () => parseCliArgs(args),
      (error) => error.code === "INVALID_ARGUMENT",
    );
  }
  assert.throws(
    () =>
      parseCliArgs([
        "measure",
        "--batch",
        "--recursive",
        "--no-recursive",
        "x",
      ]),
    /cannot be combined/,
  );
  assert.throws(
    () => parseCliArgs(["measure", "--batch", "-"]),
    /Positional batch does not accept stdin/,
  );
  assert.doesNotThrow(() =>
    parseCliArgs([
      "measure",
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
    ["measure", "--sections", "--page-limit", "1"],
    "# Analysis\n\nBody.\n\n<!-- pagebreak -->\n\nMore.",
  );
  assert.equal(human.code, 0, human.stderr);
  assert.match(
    human.stdout,
    /Section 0 \(preamble\): 0 pages; 0 visual lines; 0 counted lines/,
  );
  assert.match(human.stdout, /Section 1 \(H1 "Analysis"\): pages 1,2 \(2\)/);
  assert.match(human.stdout, /beyond limit 1: 2/);

  const json = await runInProcess(
    ["measure", "--sections", "--json"],
    "# JSON\n\nBody.",
  );
  assert.equal(
    JSON.parse(json.stdout).deterministic.sections[1].source,
    "deterministic",
  );
  const regular = await runInProcess(["measure", "--json"], "# JSON\n\nBody.");
  assert.equal("sections" in JSON.parse(regular.stdout).deterministic, false);

  const batch = await runInProcess(
    ["measure", "--batch", "--input-jsonl", "--sections"],
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
        "measure",
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
      ["measure", "--batch", "--no-recursive", "docs"],
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
      ["measure", "--batch", "--include", "*.txt", "explicit.txt"],
      "",
      { cwd: temporary },
    );
    assert.equal(custom.code, 0);
    assert.equal(JSON.parse(custom.stdout).source.path, "explicit.txt");

    try {
      await symlink(join(docs, "a.md"), join(temporary, "alias.md"));
      const aliases = await runInProcess(
        ["measure", "--batch", "alias.md", "docs/a.md"],
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
  const temporary = await mkdtemp(join(tmpdir(), "agent-docx-batch-config-"));
  await mkdir(join(temporary, "docs"));
  await writeFile(join(temporary, "docs", "a.md"), "A.");
  await writeFile(join(temporary, "docs", "a.txt"), "Text.");
  try {
    for (const selector of ["docs/no-match-*.md", "docs/[broken"]) {
      const result = await runInProcess(["measure", "--batch", selector], "", {
        cwd: temporary,
      });
      assert.equal(result.code, 2);
      assert.equal(result.stdout, "");
      assert.match(result.stderr, /INVALID_ARGUMENT/);
    }
    const filtered = await runInProcess(
      ["measure", "--batch", "--exclude", "*.md", "docs"],
      "",
      { cwd: temporary },
    );
    assert.equal(filtered.code, 2);
    assert.equal(filtered.stdout, "");

    const missing = await runInProcess(
      ["measure", "--batch", "missing.md", "docs/a.md"],
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
      ["measure", "--batch", "--config", configPath, "docs"],
      "",
      { cwd: temporary },
    );
    assert.equal(JSON.parse(configured.stdout).source.path, "docs/a.txt");
    const replaced = await runInProcess(
      [
        "measure",
        "--batch",
        "--config",
        configPath,
        "--include",
        "*.md",
        "docs",
      ],
      "",
      { cwd: temporary },
    );
    assert.equal(JSON.parse(replaced.stdout).source.path, "docs/a.md");

    await writeFile(configPath, JSON.stringify({ batch: { include: [] } }));
    const invalidConfig = await runInProcess(
      ["measure", "--batch", "--config", configPath, "docs"],
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
  const temporary = await mkdtemp(join(tmpdir(), "agent-docx-watch-"));
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
      ["measure", "--watch", "--jsonl", "watch.md"],
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

test("end-to-end legal project fixture exports clean and native-redline DOCX", async () => {
  const directory = await mkdtemp(join(tmpdir(), "agent-docx-fixture-"));
  const fixture = join(root, "test", "fixtures", "agent-project");
  const manifest = join(directory, "agent-docx.json");
  try {
    for (const file of ["motion.md", "metadata.json", "chrome.json"])
      await writeFile(
        join(directory, file),
        await readFile(join(fixture, file)),
      );

    const initialized = await runInProcess(
      [
        "project",
        "init",
        "--project",
        "agent-docx.json",
        "--document",
        "motion",
        "--source",
        "motion.md",
        "--profile",
        "cand-civil",
        "--metadata",
        "metadata.json",
        "--chrome",
        "chrome.json",
        "--filing-kind",
        "motion-document",
        "--rule-pack",
        "cand-civil@2026-05-01",
        "--json",
      ],
      "",
      { cwd: directory },
    );
    assert.equal(initialized.code, 0, initialized.stderr);

    const first = await runInProcess(
      [
        "revision",
        "checkpoint",
        "--project",
        "agent-docx.json",
        "--document",
        "motion",
        "--author",
        "Drafter",
        "--message",
        "Fixture initial draft",
        "--json",
      ],
      "",
      { cwd: directory },
    );
    assert.equal(first.code, 0, first.stderr);
    const firstRevision = JSON.parse(first.stdout).revision.id;

    const validation = await runInProcess(
      [
        "validate",
        "--project",
        "agent-docx.json",
        "--document",
        "motion",
        "--json",
      ],
      "",
      { cwd: directory },
    );
    assert.equal(validation.code, 0, validation.stderr);
    assert.match(JSON.parse(validation.stdout).status, /^(pass|fail|unknown)$/);

    const cleanPath = join(directory, "motion-clean.docx");
    const clean = await runInProcess(
      [
        "export",
        "--project",
        "agent-docx.json",
        "--document",
        "motion",
        "--revision",
        firstRevision,
        "--mode",
        "clean",
        "--output",
        cleanPath,
        "--json",
      ],
      "",
      { cwd: directory },
    );
    assert.equal(clean.code, 0, clean.stderr);
    assert.equal((await readFile(cleanPath)).subarray(0, 2).toString(), "PK");

    const marked = await readFile(join(directory, "motion.md"), "utf8");
    await writeFile(
      join(directory, "motion.md"),
      marked.replace(
        "the record establishes the required elements.",
        "the record establishes the required elements and no narrower remedy will cure the injury.",
      ),
    );
    const second = await runInProcess(
      [
        "revision",
        "checkpoint",
        "--project",
        "agent-docx.json",
        "--document",
        "motion",
        "--base",
        firstRevision,
        "--author",
        "Drafter",
        "--message",
        "Fixture revised draft",
        "--json",
      ],
      "",
      { cwd: directory },
    );
    assert.equal(second.code, 0, second.stderr);
    const secondRevision = JSON.parse(second.stdout).revision.id;

    const redlinePath = join(directory, "motion-redline.docx");
    const redline = await runInProcess(
      [
        "export",
        "--project",
        "agent-docx.json",
        "--document",
        "motion",
        "--revision",
        secondRevision,
        "--base",
        firstRevision,
        "--mode",
        "redline",
        "--output",
        redlinePath,
        "--json",
      ],
      "",
      { cwd: directory },
    );
    assert.equal(redline.code, 0, redline.stderr);
    assert.equal((await readFile(redlinePath)).subarray(0, 2).toString(), "PK");

    const agentRequest = {
      schemaVersion: 1,
      id: "fixture-measure",
      action: "document.measure",
      project: manifest,
      params: { documentId: "motion", revision: "HEAD" },
    };
    assert.equal(
      validateAgentRequest(agentRequest),
      true,
      JSON.stringify(validateAgentRequest.errors),
    );
    const agent = await runInProcess(
      ["agent", "--input-jsonl"],
      `${JSON.stringify(agentRequest)}\n`,
      { cwd: directory },
    );
    assert.equal(agent.code, 0, agent.stderr);
    const agentResult = JSON.parse(agent.stdout);
    assert.equal(agentResult.kind, "result", agent.stdout);
    assert.equal(agentResult.revision, secondRevision);
    assert.equal(agentResult.value.documentId, "motion");
    assert.equal(
      validateAgentResponse(agentResult),
      true,
      JSON.stringify(validateAgentResponse.errors),
    );

    const workingDocumentAgent = await runInProcess(
      ["agent", "--input-jsonl"],
      `${JSON.stringify({
        schemaVersion: 1,
        id: "fixture-working-document",
        action: "document.get",
        project: manifest,
        params: { documentId: "motion" },
      })}\n`,
      { cwd: directory },
    );
    assert.equal(workingDocumentAgent.code, 0, workingDocumentAgent.stderr);
    const workingDocumentResult = JSON.parse(workingDocumentAgent.stdout);
    assert.equal(workingDocumentResult.revision, null);
    assert.equal(
      validateAgentResponse(workingDocumentResult),
      true,
      JSON.stringify(validateAgentResponse.errors),
    );
    const watchedSource = await readFile(join(directory, "motion.md"), "utf8");
    const capture = memoryRuntime("", {
      cwd: directory,
      onceSignal(signal, listener) {
        if (signal !== "SIGINT") return;
        setTimeout(
          () =>
            void writeFile(
              join(directory, "motion.md"),
              watchedSource.replace(
                "the Court should grant the motion.",
                "the Court should grant the requested motion.",
              ),
            ),
          10,
        );
        setTimeout(listener, 600);
      },
    });
    const watchCode = await runCli(
      [
        "agent",
        "--watch",
        "--project",
        "agent-docx.json",
        "--document",
        "motion",
        "--jsonl",
      ],
      capture.runtime,
    );
    assert.equal(watchCode, 0);
    const records = capture.stdout().trim().split("\n").map(JSON.parse);
    assert.deepEqual(
      records.map(({ kind }) => kind),
      ["ready", "result", "end"],
    );
    assert.equal(records[1].action, "document.measure");
    assert.equal(records[1].revision, null);
    for (const record of records)
      assert.equal(
        validateAgentStream(record),
        true,
        JSON.stringify(validateAgentStream.errors),
      );
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});

test("agent protocol closes stateless, result, error, and fatal envelopes", () => {
  const inspectWithProject = {
    schemaVersion: 1,
    action: "docx.inspect",
    project: "agent-docx.json",
    params: { input: "template.docx" },
  };
  assert.equal(validateAgentRequest(inspectWithProject), false);

  const malformedBlockReview = {
    schemaVersion: 1,
    action: "review.add",
    params: {
      documentId: "motion",
      revision: "HEAD",
      blockId: "b_00000000-0000-0000-0000-000000000000",
      author: { name: "Reviewer" },
      message: "Review this block",
    },
  };
  assert.equal(validateAgentRequest(malformedBlockReview), false);
  assert.throws(() => parseAgentRequest(malformedBlockReview), /block ID/);

  const fatal = {
    schemaVersion: 1,
    kind: "fatal",
    sequence: 1,
    requestId: null,
    action: null,
    project: null,
    documentId: null,
    revision: null,
    error: { code: "INPUT_NOT_UTF8", message: "Input is not valid UTF-8" },
  };
  assert.equal(
    validateAgentResponse(fatal),
    true,
    JSON.stringify(validateAgentResponse.errors),
  );
  assert.equal(validateAgentResponse({ ...fatal, value: null }), false);

  const invalidMeasureResult = {
    schemaVersion: 1,
    kind: "result",
    sequence: 1,
    requestId: "measure",
    action: "document.measure",
    project: "agent-docx.json",
    documentId: "motion",
    revision: null,
    value: {},
  };
  assert.equal(validateAgentResponse(invalidMeasureResult), false);
});

test("agent JSONL shares sequence numbers with a UTF-8 fatal record", async () => {
  const first = Buffer.from(
    `${JSON.stringify({
      schemaVersion: 1,
      id: "first",
      action: "project.get",
      params: { extra: true },
    })}\n`,
  );
  const result = await runInProcess(["agent", "--input-jsonl"], "", {
    readStdin: async () => {
      throw new Error("agent JSONL must use streaming stdin");
    },
    readStdinChunks: async function* () {
      yield first;
      yield Uint8Array.from([0xff]);
    },
  });
  assert.equal(result.code, 1);
  const stdout = JSON.parse(result.stdout);
  const stderr = JSON.parse(result.stderr);
  assert.equal(stdout.kind, "error");
  assert.equal(stdout.sequence, 1);
  assert.equal(
    validateAgentResponse(stdout),
    true,
    JSON.stringify(validateAgentResponse.errors),
  );
  assert.equal(stderr.kind, "fatal");
  assert.equal(stderr.sequence, 2);
  assert.equal(
    validateAgentResponse(stderr),
    true,
    JSON.stringify(validateAgentResponse.errors),
  );
});

test("inspect-only import rejects stateful CLI options and strips binary payloads", async () => {
  const result = await runInProcess([
    "import",
    "input.docx",
    "--inspect-only",
    "--project",
    "agent-docx.json",
  ]);
  assert.equal(result.code, 2);
  assert.match(result.stderr, /INVALID_ARGUMENT/);

  const serialized = serializeAgentValue({
    bytes: Uint8Array.from([0x50, 0x4b]),
    attachments: {
      manifestSha256:
        "sha256:aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa",
      files: {
        "exhibit.pdf": {
          bytes: Uint8Array.from([1, 2, 3]),
          mediaType: "application/pdf",
        },
      },
    },
    artifact: {
      byteLength: 2,
      sha256:
        "sha256:bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb",
    },
  });
  assert.deepEqual(serialized, {
    artifact: {
      byteLength: 2,
      sha256:
        "sha256:bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb",
    },
  });
});
