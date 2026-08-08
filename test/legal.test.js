import assert from "node:assert/strict";
import test from "node:test";
import {
  insertMissingBlockMarkers,
  parseLegalMarkdown,
} from "../dist/legal/parse.js";
import { lowerLegalDocument } from "../dist/legal/lower.js";
import { metadata } from "./helpers.js";

test("legal parser materializes stable block markers and source-mapped runs", () => {
  const source = `# Argument

The **first** point has a [link](https://example.test) and a footnote.[^n]

[^n]: Supporting authority.`;
  const marked = insertMissingBlockMarkers(source, {
    documentId: "motion",
    metadata,
  });
  assert.match(marked, /<!-- agent-docx:block id="b_[0-9a-f-]+" -->/);

  const { document, missingMarkers } = parseLegalMarkdown(marked, {
    documentId: "motion",
    metadata,
    requireMarkers: true,
  });
  assert.equal(missingMarkers.length, 0);
  assert.equal(document.blocks.length, 2);
  assert.equal(document.footnotes.length, 1);
  assert.equal(document.blocks[0].kind, "heading");
  assert.equal(document.blocks[1].kind, "paragraph");
  const paragraph = document.blocks[1];
  assert.equal(paragraph.kind, "paragraph");
  assert.equal(
    paragraph.runs.some((run) => run.bold),
    true,
  );
  assert.equal(
    paragraph.runs.some((run) => run.link?.target === "https://example.test"),
    true,
  );
  assert.equal(
    paragraph.runs.some((run) => run.footnoteId === "n"),
    true,
  );
  const lowered = lowerLegalDocument(document);
  assert.equal(lowered.blocks.length, 2);
  assert.equal(lowered.footnotes.get("n")?.blocks.length, 1);
});

test("legal parser rejects unsafe ordinary links", () => {
  assert.throws(
    () =>
      parseLegalMarkdown("[local](file:///tmp/brief.md)", {
        documentId: "motion",
        metadata,
      }),
    (error) => error?.code === "REFERENCE_INVALID",
  );
});
