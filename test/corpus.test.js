import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { readFile } from "node:fs/promises";
import test from "node:test";
import { measureMarkdown } from "../dist/index.js";

const manifest = JSON.parse(
  await readFile(new URL("./corpus-manifest.json", import.meta.url), "utf8"),
);

const sha256 = (value) => createHash("sha256").update(value).digest("hex");

for (const document of manifest.documents) {
  test(`real Markdown corpus: ${document.id}`, async () => {
    const markdown = await readFile(
      new URL(`./fixtures/corpus/${document.file}`, import.meta.url),
      "utf8",
    );
    assert.equal(
      sha256(markdown),
      document.markdownSha256,
      "fixture content changed without updating its deterministic golden",
    );

    const result = await measureMarkdown(markdown, {
      paragraphDiagnostics: true,
    });
    const deterministic = result.deterministic;
    assert.deepEqual(
      {
        pageCount: deterministic.pageCount,
        equivalentPages: deterministic.equivalentPages,
        totalVisualLines: deterministic.totalVisualLines,
        paragraphCount: deterministic.paragraphs.length,
        lastPage: deterministic.lastPage,
        warningCodes: deterministic.warnings.map(({ code }) => code),
      },
      document.expected,
    );
  });
}
