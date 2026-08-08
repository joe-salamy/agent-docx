import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";
import {
  createProject,
  compileMarkdown,
  builtInRulePacks,
} from "../dist/index.js";
import { parseLegalMarkdown } from "../dist/legal/parse.js";
import { validateLegalDocument } from "../dist/legal/rules.js";
import { lowerLegalDocument } from "../dist/legal/lower.js";
import { visibleBlock } from "../dist/legal/visible-text.js";
import { normalizeMarkdown } from "../dist/markdown.js";
import { readDocxParts, repackDocxParts } from "../dist/docx/package.js";
import { attachmentInventory } from "../dist/docx/attachments.js";
import { metadata } from "./helpers.js";

const sourceHash = `sha256:${"0".repeat(64)}`;
const packWith = (checks, overrides = {}) => ({
  id: "local-rules@2026-01-01",
  sourceUrl: "https://example.test/local-rules",
  effectiveDate: "2026-01-01",
  sourceSha256: sourceHash,
  sourceExcerpt: "local-rules.txt",
  checks,
  unmodeledProvisions: [],
  ...overrides,
});
const check = (id, kind, params) => ({
  id,
  kind,
  citation: `Local Rule ${id}`,
  predicate: `predicate for ${id}`,
  params,
});

test("footnote references are validated in headings, list items, and table cells", () => {
  const sources = [
    "# Heading[^missing]\n\nBody.\n",
    "1. Item[^missing]\n\nBody.\n",
    "| Ref[^missing] |\n| --- |\n| cell |\n",
  ];
  for (const source of sources)
    assert.throws(
      () =>
        parseLegalMarkdown(source, {
          documentId: "motion",
          metadata,
        }),
      (error) => error?.code === "REFERENCE_INVALID",
      `expected missing-footnote rejection for: ${source}`,
    );
});

test("escaped footnote-looking text stays literal", () => {
  const source = "Literal \\[^id] stays text.\n";
  const parsed = parseLegalMarkdown(source, {
    documentId: "motion",
    metadata,
  });
  assert.equal(parsed.document.footnotes.length, 0);
  const paragraph = parsed.document.blocks.find(
    (block) => block.kind === "paragraph",
  );
  assert.equal(paragraph?.kind, "paragraph");
  assert.ok(
    paragraph.runs.some((run) => run.text.includes("[^id]")),
    "escaped footnote text must remain literal",
  );
});

test("block marker at EOF without trailing newline is recognized as an orphan", () => {
  assert.throws(
    () =>
      parseLegalMarkdown(
        '<!-- agent-docx:block id="b_12345678-1234-4123-8123-123456789abc" -->',
        { documentId: "motion", metadata, requireMarkers: true },
      ),
    (error) => error?.code === "REFERENCE_INVALID",
  );
});

test("deeply nested lists fail with a bounded error, not a stack overflow", () => {
  let source = "- bottom\n";
  for (let depth = 0; depth < 120; depth++)
    source = `- outer\n${"  ".repeat(1)}${source}`;
  // Build 150 nested list levels: each level adds two-space indentation.
  let nested = "text";
  for (let level = 0; level < 150; level++) nested = `- ${nested}`;
  const indented = nested
    .split("\n")
    .map((line, index) => `${"  ".repeat(index)}${line}`)
    .join("\n");
  assert.throws(
    () => parseLegalMarkdown(indented, { documentId: "motion", metadata }),
    (error) =>
      error?.code === "UNSUPPORTED_MARKDOWN" &&
      /nesting/i.test(error?.message ?? ""),
  );
  void source;
});

test("public parser rejects invalid document ids", () => {
  assert.throws(
    () =>
      parseLegalMarkdown("# Motion\n", { documentId: "Bad ID/..", metadata }),
    (error) => error?.code === "INVALID_ARGUMENT",
  );
});

test("markdown normalization rejects XML-illegal code points with a source position", () => {
  for (const source of ["Bad\x01char.\n", "Lone \ud800 surrogate.\n"])
    assert.throws(
      () => normalizeMarkdown(source),
      (error) =>
        error?.code === "UNSUPPORTED_MARKDOWN" &&
        /XML-1\.0-illegal/i.test(error?.message ?? ""),
    );
});

test("measurement normalization rejects unsafe link targets", () => {
  for (const target of [
    "javascript:alert(1)",
    "file:///tmp/x",
    "data:text/plain,x",
  ])
    assert.throws(
      () => normalizeMarkdown(`[x](${target})\n`),
      (error) => error?.code === "REFERENCE_INVALID",
    );
});

test("ragged table rows are padded to a rectangle", () => {
  const source = "| A |\n| - |\n| one | two |\n";
  const { document } = parseLegalMarkdown(source, {
    documentId: "motion",
    metadata,
  });
  const table = document.blocks.find((block) => block.kind === "table");
  assert.equal(table?.kind, "table");
  assert.equal(table.rows.length, 2);
  assert.equal(table.rows[0].length, 2, "header row padded");
  assert.equal(table.rows[1].length, 2);
  assert.equal(table.rows[1][1].paragraphs[0]?.runs[0]?.text, "two");
  const lowered = lowerLegalDocument(document);
  const loweredTable = lowered.blocks.find((block) => block.kind === "table");
  assert.equal(loweredTable?.kind, "table");
  assert.equal(loweredTable.rows[0].length, 2);
});

test("footnote labels match across NFC-equivalent forms", () => {
  const source = "Ref.[^e\u0301]\n\n[^é]: Definition.\n";
  const parsed = parseLegalMarkdown(source, {
    documentId: "motion",
    metadata,
  });
  assert.equal(parsed.document.footnotes.length, 1);
});

test("hard breaks are represented in normalized segments and visible text", () => {
  const source = "First.  \nSecond.\n";
  const { document } = parseLegalMarkdown(source, {
    documentId: "motion",
    metadata,
  });
  const paragraph = document.blocks.find((block) => block.kind === "paragraph");
  assert.equal(paragraph?.kind, "paragraph");
  const lowered = lowerLegalDocument(document);
  const loweredParagraph = lowered.blocks.find(
    (block) => block.kind === "paragraph",
  );
  assert.equal(loweredParagraph?.kind, "paragraph");
  assert.equal(loweredParagraph.normalizedText, "First.\nSecond.");
  assert.ok(
    paragraph.segments.some(
      (segment) => segment.normalizedEnd - segment.normalizedStart === 1,
    ),
    "hard break must be represented in the source segments",
  );
});

test("signature visible text matches the rendered counsel name", () => {
  const withCounsel = {
    ...metadata,
    counsel: [{ id: "counsel", name: "Alice T. Counsel" }],
  };
  const { document } = parseLegalMarkdown(
    '::signature{counsel="counsel"}\n\nBody.\n',
    { documentId: "motion", metadata: withCounsel },
  );
  const signature = document.blocks.find((block) => block.kind === "signature");
  assert.equal(signature?.kind, "signature");
  assert.equal(visibleBlock(signature, withCounsel), "Alice T. Counsel");
});

test("exhibit text is counted once, and unreferenced footnotes are not counted", async () => {
  const exhibitWords = Array.from(
    { length: 60 },
    (_, index) => `word${index}`,
  ).join(" ");
  const source = [
    "# Motion",
    "",
    `:::exhibit{id="ex1" label="Exhibit A" source="exhibit.txt"}`,
    exhibitWords,
    ":::",
    "",
    "Body text.[^used]",
    "",
    "[^used]: Used definition.",
    "[^unused]: " +
      Array.from({ length: 500 }, (_, index) => `junk${index}`).join(" "),
    "",
  ].join("\n");
  const { document } = parseLegalMarkdown(source, {
    documentId: "motion",
    metadata,
    assets: {
      "exhibit.txt": {
        bytes: new TextEncoder().encode(exhibitWords),
        mediaType: "text/plain",
      },
    },
  });
  const wordPack = packWith([
    check("length", "length-alternative", {
      byFilingKind: {
        default: { words: 80, complianceCertificateRequired: false },
      },
    }),
  ]);
  const validation = validateLegalDocument(document, {
    customPacks: [wordPack],
  });
  const length = validation.findings.find(
    (finding) => finding.checkId === "length",
  );
  assert.equal(
    length?.status,
    "pass",
    `exhibit (60 words) + body must count under 80 words, got: ${length?.status}`,
  );
});

test("custom certificate checks require a rendered certificate block", async () => {
  const source = "# Motion\n\nBody.\n";
  const { document } = parseLegalMarkdown(source, {
    documentId: "motion",
    metadata,
  });
  const pack = packWith([
    check("cert", "required-metadata", {
      fields: ["caseName", "docketNumber", "documentTitle"],
      requireComplianceCertificate: true,
    }),
  ]);
  const validation = validateLegalDocument(document, { customPacks: [pack] });
  assert.equal(
    validation.findings.find((finding) => finding.checkId === "cert")?.status,
    "fail",
    "metadata-only certificate must not satisfy the requirement",
  );
});

test("rule pack and check ids must be unique across custom packs", async () => {
  const { document } = parseLegalMarkdown("# Motion\n\nBody.\n", {
    documentId: "motion",
    metadata,
  });
  const packA = packWith([
    check("length", "required-block", { kinds: ["paragraph"] }),
  ]);
  const packB = {
    ...packWith([check("other", "required-block", { kinds: ["paragraph"] })]),
    id: "local-rules@2026-01-01",
  };
  assert.throws(
    () => validateLegalDocument(document, { customPacks: [packA, packB] }),
    (error) =>
      error?.code === "RULE_PACK_INVALID" &&
      /Duplicate rule pack id/.test(error.message),
  );
  const packC = packWith(
    [check("length", "required-block", { kinds: ["paragraph"] })],
    {
      id: "other-rules@2026-01-01",
    },
  );
  assert.throws(
    () => validateLegalDocument(document, { customPacks: [packA, packC] }),
    (error) =>
      error?.code === "RULE_PACK_INVALID" &&
      /Duplicate rule check id/.test(error.message),
  );
});

test("rule pack effective dates are real calendar dates", () => {
  const bad = packWith(
    [check("block", "required-block", { kinds: ["paragraph"] })],
    {
      effectiveDate: "2026-99-99",
    },
  );
  assert.throws(
    () =>
      validateLegalDocument(
        parseLegalMarkdown("# M\n\nB.\n", { documentId: "motion", metadata })
          .document,
        { customPacks: [bad] },
      ),
    (error) => error?.code === "RULE_PACK_INVALID",
  );
});

test("compilation with an external hyperlink is byte-deterministic", async () => {
  const options = {
    documentId: "motion",
    profile: "us-district-conventional",
    metadata,
  };
  const source = "# Motion\n\nSee [authority](https://example.test/case).\n";
  const first = await compileMarkdown(source, options);
  const second = await compileMarkdown(source, options);
  assert.deepEqual(first.bytes, second.bytes);
  assert.equal(
    createHash("sha256").update(first.bytes).digest("hex"),
    createHash("sha256").update(second.bytes).digest("hex"),
  );
});

test("ZIP import rejects a CRC mismatch", async () => {
  const compiled = await compileMarkdown("# Motion\n\nBody.\n", {
    documentId: "motion",
    profile: "us-district-conventional",
    metadata,
  });
  const parts = await readDocxParts(compiled.bytes);
  const repacked = Buffer.from(repackDocxParts(new Map(parts)));
  // Corrupt both CRC-32 fields (local header + central directory) for the
  // first entry; the deflate stream stays intact, so only CRC can catch it.
  const localCrc = repacked.indexOf(Buffer.from([0x50, 0x4b, 0x03, 0x04])) + 14;
  repacked.writeUInt32LE(0, localCrc);
  const centralCrc =
    repacked.indexOf(Buffer.from([0x50, 0x4b, 0x01, 0x02])) + 16;
  repacked.writeUInt32LE(0, centralCrc);
  await assert.rejects(
    () => readDocxParts(repacked),
    (error) => error?.code === "DOCX_INVALID",
  );
});

test("page numbering start is emitted only on the first section", async () => {
  const compiled = await compileMarkdown(
    '# Motion\n\nFirst section.\n\n::sectionbreak{kind="next-page"}\n\nSecond section.\n',
    {
      documentId: "motion",
      profile: "us-district-conventional",
      metadata,
      chrome: {
        pageNumber: {
          story: "footer",
          alignment: "center",
          format: "decimal",
          start: 1,
        },
      },
    },
  );
  const parts = await readDocxParts(compiled.bytes);
  const xml = new TextDecoder().decode(parts.get("word/document.xml"));
  const started = xml.match(/<w:pgNumType[^>]*w:start="1"/g) ?? [];
  assert.equal(
    started.length,
    1,
    `page-number start must be emitted once, found ${started.length}`,
  );
});

test("attachment inventory keys parts by full path, not basename", () => {
  const bytes = (value) => new TextEncoder().encode(value);
  const parts = new Map([
    ["word/media/a/foo.png", bytes("a")],
    ["word/media/b/foo.png", bytes("b")],
  ]);
  const inventory = attachmentInventory(parts);
  const keys = Object.keys(inventory);
  assert.ok(keys.includes("word/media/a/foo.png"));
  assert.ok(keys.includes("word/media/b/foo.png"));
});

test("delete-block redlines round-trip through import with resolution", async () => {
  const root = await mkdtemp(join(tmpdir(), "agent-docx-delete-block-"));
  try {
    const manifestPath = join(root, "agent-docx.json");
    const sourcePath = join(root, "motion.md");
    await writeFile(
      sourcePath,
      "# Motion\n\nFirst paragraph.\n\nSecond paragraph.\n",
    );
    const project = await createProject(manifestPath, {
      documentId: "motion",
      source: "motion.md",
      profile: "us-district-conventional",
      metadata,
    });
    const base = await project.checkpoint("motion", {
      baseRevision: null,
      author: { name: "Drafter" },
      message: "Initial draft",
    });
    await writeFile(sourcePath, "# Motion\n\nSecond paragraph.\n");
    const head = await project.checkpoint("motion", {
      baseRevision: base.revision.id,
      author: { name: "Drafter" },
      message: "Delete first paragraph",
    });
    const exported = await project.exportDocx("motion", {
      revision: head.revision.id,
      mode: "redline",
      baseRevision: base.revision.id,
      output: join(root, "motion-redline.docx"),
    });
    assert.ok(exported.bytes.byteLength > 0);
    const imported = await project.importRedline({
      documentId: "motion",
      input: join(root, "motion-redline.docx"),
      author: { name: "Reviewer" },
      message: "Review redline",
    });
    assert.equal(imported.headRevision, head.revision.id);
    assert.equal(imported.baseRevision, base.revision.id);
    assert.ok(
      imported.changeSet.changes.some(
        (change) => change.kind === "delete-block",
      ),
      "imported change set must include the delete-block change",
    );
  } finally {
    await rm(root, { recursive: true, force: true });
  }
});

test("input reads enforce byte caps and reject symlinks", async () => {
  const root = await mkdtemp(join(tmpdir(), "agent-docx-input-cap-"));
  try {
    const small = join(root, "small.md");
    const large = join(root, "large.md");
    await writeFile(small, "# Small\n");
    await writeFile(large, "x".repeat(4096));
    const { readInputFile } = await import("../dist/input.js");
    assert.deepEqual(
      await readInputFile(small, "Small"),
      new TextEncoder().encode("# Small\n"),
    );
    await assert.rejects(
      () => readInputFile(large, "Large", 1024),
      (error) => error?.code === "INPUT_TOO_LARGE",
    );
    if (process.platform !== "win32") {
      const link = join(root, "linked.md");
      const { symlink } = await import("node:fs/promises");
      await symlink(small, link);
      await assert.rejects(
        () => readInputFile(link, "Linked"),
        (error) => error?.code === "INPUT_NOT_FOUND",
      );
    }
  } finally {
    await rm(root, { recursive: true, force: true });
  }
});

test("built-in rule pack binding still holds after validation hardening", async () => {
  const pack = builtInRulePacks["cand-civil@2026-05-01"];
  const compiled = await compileMarkdown("# Motion\n\nRequested relief.\n", {
    documentId: "motion",
    profile: "cand-civil",
    filingKind: "motion-document",
    rulePack: "cand-civil@2026-05-01",
    metadata,
  });
  assert.equal(compiled.validation.status, "fail");
  assert.ok(
    compiled.validation.findings.some(
      (finding) => finding.checkId === "cand.footer",
    ),
  );
  assert.equal(pack.id, "cand-civil@2026-05-01");
});
