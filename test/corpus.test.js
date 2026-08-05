import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { readFile } from "node:fs/promises";
import test from "node:test";
import { runSubprocess } from "./helpers.js";

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
const fixturePath = (document) =>
  `test/fixtures/${document.fixtureDirectory}/${document.file}`;

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

    const invocation = await runSubprocess([
      "measure",
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
  { timeout: 600_000, skip: process.env.AGENT_DOCX_TEST_WORD !== "1" },
  async (t) => {
    const invocation = await runSubprocess([
      "measure",
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
