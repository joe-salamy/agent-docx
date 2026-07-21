import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { createHash } from "node:crypto";
import { readFile } from "node:fs/promises";
import test from "node:test";
import { fileURLToPath } from "node:url";

const manifest = JSON.parse(
  await readFile(new URL("./corpus-manifest.json", import.meta.url), "utf8"),
);
const blindManifest = JSON.parse(
  await readFile(
    new URL("./blind-corpus-manifest.json", import.meta.url),
    "utf8",
  ),
);
const documents = [
  ...manifest.documents.map((document) => ({
    ...document,
    fixtureDirectory: "corpus",
    requireAllPageLines: true,
  })),
  ...blindManifest.documents.map((document) => ({
    ...document,
    fixtureDirectory: "blind-corpus",
    requireAllPageLines: false,
  })),
];
const root = fileURLToPath(new URL("..", import.meta.url));
const cli = fileURLToPath(new URL("../dist/cli.js", import.meta.url));
const fixturePath = (document) =>
  `test/fixtures/${document.fixtureDirectory}/${document.file}`;

function runCli(args) {
  const { promise, resolve, reject } = Promise.withResolvers();
  const child = spawn(process.execPath, [cli, ...args], {
    cwd: root,
    stdio: ["ignore", "pipe", "pipe"],
  });
  let stdout = "";
  let stderr = "";
  child.stdout.setEncoding("utf8").on("data", (chunk) => (stdout += chunk));
  child.stderr.setEncoding("utf8").on("data", (chunk) => (stderr += chunk));
  child.once("error", reject);
  child.once("close", (code) => resolve({ code, stdout, stderr }));
  return promise;
}

function deterministicGolden(deterministic) {
  return {
    pageCount: deterministic.pageCount,
    equivalentPages: deterministic.equivalentPages,
    totalVisualLines: deterministic.totalVisualLines,
    visualLinesByPage: deterministic.visualLinesByPage,
    paragraphCount: deterministic.paragraphs.length,
    lastPage: deterministic.lastPage,
    warningCodes: deterministic.warnings.map(({ code }) => code),
  };
}

const sha256 = (value) => createHash("sha256").update(value).digest("hex");

for (const document of documents) {
  test(`real Markdown corpus: ${document.id}`, async () => {
    const markdown = await readFile(
      new URL(
        `./fixtures/${document.fixtureDirectory}/${document.file}`,
        import.meta.url,
      ),
      "utf8",
    );
    assert.equal(
      sha256(markdown),
      document.markdownSha256,
      "fixture content changed without updating its deterministic golden",
    );

    const invocation = await runCli([
      fixturePath(document),
      "--paragraphs",
      "--json",
    ]);
    assert.equal(invocation.code, 0, invocation.stderr);
    assert.equal(invocation.stderr, "");
    const result = JSON.parse(invocation.stdout);
    assert.deepEqual(
      deterministicGolden(result.deterministic),
      document.expected,
    );
  });
}

test(
  "real Markdown corpus: CLI page and last-page lines agree exactly with Word",
  { timeout: 600_000 },
  async (t) => {
    const invocation = await runCli([
      "--batch",
      ...documents.map(fixturePath),
      "--paragraphs",
      "--renderer",
      "word",
    ]);
    const records = invocation.stdout
      .trim()
      .split("\n")
      .filter(Boolean)
      .map(JSON.parse);
    const unavailable =
      records.length > 0 &&
      records.every(
        (record) =>
          record.kind === "error" &&
          ["WORD_NOT_FOUND", "WORD_WSL_BRIDGE_UNAVAILABLE"].includes(
            record.error?.code,
          ),
      );
    if (unavailable) {
      t.skip("desktop Microsoft Word is unavailable");
      return;
    }
    assert.equal(invocation.code, 0, invocation.stderr || invocation.stdout);
    assert.equal(records.length, documents.length);
    for (const [index, record] of records.entries()) {
      assert.equal(record.kind, "result");
      const document = documents[index];
      const deterministic = record.measurement.deterministic;
      const word = record.measurement.renderers.word.value;
      assert.deepEqual(deterministicGolden(deterministic), document.expected);
      assert.equal(word.pageCount, deterministic.pageCount);
      if (document.requireAllPageLines) {
        assert.equal(word.totalBodyLines, deterministic.totalVisualLines);
        assert.deepEqual(word.bodyLinesByPage, deterministic.visualLinesByPage);
      }
      assert.equal(
        word.bodyLinesOnLastPage,
        deterministic.lastPage.visualLines,
      );
    }
  },
);
