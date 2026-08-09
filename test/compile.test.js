import assert from "node:assert/strict";
import { mkdtemp, mkdir, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";
import {
  AgentDocxError,
  builtInProfiles,
  compileMarkdown,
  generateDocx,
  generateRedlineDocx,
  inspectDocx,
  lowerLegalDocument,
  parseLegalMarkdown,
} from "../dist/index.js";
import { measureNormalizedDocument } from "../dist/renderers/index.js";
import { loadFonts } from "../dist/resolve.js";
import { tableColumnWidths } from "../dist/layout/table.js";
import { inspectDocxMaterial } from "../dist/docx/import.js";
import { readDocxParts, repackDocxParts } from "../dist/docx/package.js";
import { metadata } from "./helpers.js";

test("standalone compilation parses legal Markdown and emits stable block bookmarks", async () => {
  const compiled = await compileMarkdown(
    "# Motion\n\nThe requested relief follows.\n",
    {
      documentId: "motion",
      profile: "us-district-conventional",
      metadata,
    },
  );
  assert.ok(compiled.bytes.byteLength > 1000);
  assert.equal(compiled.artifact.revision, null);
  assert.equal(compiled.validation.status, "pass");
  assert.equal(compiled.blocks.length, 2);
  for (const block of compiled.blocks)
    assert.match(block.bookmark, /^adx_[0-9a-f]{32}$/);
});

test("authority annotations survive compilation into the semantic manifest", async () => {
  const compiled = await compileMarkdown(
    ':authority[Section claim]{id="42-u-s-c-1983" category="statutes" short="42 U.S.C. § 1983"}\n\nBody.\n',
    {
      documentId: "motion",
      profile: "us-district-conventional",
      metadata,
    },
  );
  const inspected = await inspectDocx(compiled.bytes);
  const recognizedAuthority = inspected.recognized.blocks.flatMap(
    (block) =>
      block.runs?.flatMap((run) => (run.authority ? [run.authority] : [])) ??
      [],
  );
  assert.ok(
    recognizedAuthority.some(
      (authority) =>
        authority.id === "42-u-s-c-1983" &&
        authority.category === "statutes" &&
        authority.short === "42 U.S.C. § 1983",
    ),
  );
});

test("ordered lists starting above one export a numbering start override", async () => {
  const parsed = parseLegalMarkdown(
    ["4. Fourth item", "5. Fifth item", "", "1. Ordinary list"].join("\n"),
    {
      documentId: "motion",
      profile: "us-district-conventional",
      metadata,
    },
  );
  const generated = await generateDocx(
    parsed.document,
    builtInProfiles["us-district-conventional"],
  );
  const parts = await readDocxParts(generated.bytes);
  const numberingXml = new TextDecoder().decode(
    parts.get("word/numbering.xml"),
  );
  assert.match(numberingXml, /<w:start w:val="4"\/>/);
  assert.ok(
    (numberingXml.match(/<w:abstractNum/g) ?? []).length >= 4,
    "numbering.xml should carry the ordinary and start-override ordered definitions",
  );
});

test("continuous section breaks preserve the deterministic page", async () => {
  const before = Array.from(
    { length: 8 },
    (_, index) => `Before ${index}`,
  ).join("\n\n");
  const after = Array.from({ length: 8 }, (_, index) => `After ${index}`).join(
    "\n\n",
  );
  const base = `${before}\n\n${after}\n`;
  const continuous = `${before}\n\n::sectionbreak{kind="continuous"}\n\n${after}\n`;
  const nextPage = `${before}\n\n::sectionbreak{kind="next-page"}\n\n${after}\n`;
  const parseOptions = {
    documentId: "motion",
    profile: "us-district-conventional",
    metadata,
  };
  const measure = (markdown) =>
    measureNormalizedDocument(
      lowerLegalDocument(parseLegalMarkdown(markdown, parseOptions).document),
      { renderer: "deterministic" },
    );
  const [withoutBreak, withContinuous, withNextPage] = await Promise.all([
    measure(base),
    measure(continuous),
    measure(nextPage),
  ]);
  assert.equal(
    withContinuous.deterministic.pageCount,
    withoutBreak.deterministic.pageCount,
  );
  assert.equal(
    withNextPage.deterministic.pageCount,
    withoutBreak.deterministic.pageCount + 1,
  );
});

test("DOCX export allocates table columns with the deterministic widths", async () => {
  const parsed = parseLegalMarkdown(
    [
      "| Wide header column | X |",
      "| --- | --- |",
      "| A long phrase that needs room to breathe | Y |",
      "| Short | Z |",
    ].join("\n"),
    { documentId: "motion", profile: "us-district-conventional", metadata },
  );
  const lowered = lowerLegalDocument(parsed.document);
  const table = lowered.blocks.find((block) => block.kind === "table");
  const profile = builtInProfiles["us-district-conventional"];
  const fonts = await loadFonts(undefined, profile.requestedFontFamily);
  const usableWidth =
    profile.page.widthTwips -
    profile.page.marginsTwips.left -
    profile.page.marginsTwips.right -
    profile.page.gutterTwips;
  const expected = tableColumnWidths(table, profile, fonts, usableWidth);
  const generated = await generateDocx(lowered, profile, { fonts });
  const parts = await readDocxParts(generated.bytes);
  const documentXml = new TextDecoder().decode(parts.get("word/document.xml"));
  const grid = [...documentXml.matchAll(/<w:gridCol w:w="(\d+)"/g)].map(
    (match) => Number(match[1]),
  );
  assert.deepEqual(grid, expected);
});

test("native DOCX generation accepts legal IR and its document chrome", async () => {
  const parsed = parseLegalMarkdown("A plain paragraph.\n", {
    documentId: "motion",
    profile: "us-district-conventional",
    metadata,
    chrome: { headers: { default: "{{caseName}}" } },
  });
  const generated = await generateDocx(
    parsed.document,
    builtInProfiles["us-district-conventional"],
  );
  const parts = await readDocxParts(generated.bytes);
  assert.match(
    new TextDecoder().decode(parts.get("word/header1.xml")),
    /Example v\. Example/,
  );
});

test("native redline export rejects unsupported legal blocks", async () => {
  const cases = [
    ["- First item\n", "list"],
    ["Body.[^note]\n\n[^note]: Supporting note.\n", "footnote"],
    ["| Left | Right |\n| --- | --- |\n| One | Two |\n", "table"],
  ];
  for (const [source, kind] of cases) {
    const parsed = parseLegalMarkdown(source, {
      documentId: "motion",
      profile: "us-district-conventional",
      metadata,
    });
    await assert.rejects(
      generateRedlineDocx(
        parsed.document,
        parsed.document,
        { changes: [] },
        builtInProfiles["us-district-conventional"],
      ),
      (error) =>
        error instanceof AgentDocxError &&
        error.code === "DOCX_REDLINE_UNSUPPORTED" &&
        error.message.includes(kind),
    );
  }
});

test("native redline comments preserve a code-point-safe text range", async () => {
  const parsed = parseLegalMarkdown("A😀B\n", {
    documentId: "motion",
    profile: "us-district-conventional",
    metadata,
  });
  const blockId = parsed.document.blocks[0].id;
  const head = {
    ...parsed.document,
    annotations: [
      {
        id: "a_00000000-0000-4000-8000-000000000000",
        blockId,
        range: { start: 1, end: 3 },
        author: { name: "Reviewer" },
        createdAt: "2026-01-01T00:00:00.000Z",
        message: "Review the emoji.",
        status: "open",
      },
    ],
  };
  const generated = await generateRedlineDocx(
    parsed.document,
    head,
    { changes: [] },
    builtInProfiles["us-district-conventional"],
  );
  const parts = await readDocxParts(generated.bytes);
  const documentXml = new TextDecoder().decode(parts.get("word/document.xml"));
  const a = documentXml.indexOf(">A</w:t>");
  const rangeStart = documentXml.indexOf("<w:commentRangeStart");
  const emoji = documentXml.indexOf("😀");
  const rangeEnd = documentXml.indexOf("<w:commentRangeEnd");
  const reference = documentXml.indexOf("<w:commentReference");
  assert.ok(a >= 0 && a < rangeStart);
  assert.ok(rangeStart >= 0 && rangeStart < emoji);
  assert.ok(emoji >= 0 && emoji < rangeEnd);
  assert.ok(rangeEnd >= 0 && rangeEnd < reference);
});

test("caption, signature, and certificate blocks emit borderless DOCX tables", async () => {
  const compiled = await compileMarkdown(
    [
      "::caption",
      "",
      '::signature{counsel="counsel-1"}',
      "",
      '::certificate{id="certificate-1"}',
    ].join("\n"),
    {
      documentId: "motion",
      profile: "us-district-conventional",
      metadata: {
        ...metadata,
        counsel: [
          {
            id: "counsel-1",
            name: "A. Counsel",
            addressLines: [],
            phone: "",
            email: "",
          },
        ],
        certificates: [
          {
            id: "certificate-1",
            kind: "compliance",
            basis: "words",
            signerCounselId: "counsel-1",
          },
        ],
      },
    },
  );
  const parts = await readDocxParts(compiled.bytes);
  const documentXml = new TextDecoder().decode(parts.get("word/document.xml"));
  assert.equal((documentXml.match(/<w:tbl>/g) ?? []).length, 3);
  assert.match(documentXml, /w:val="none"/);
});

test("generated DOCX is boundedly inspectable without Office", async () => {
  const compiled = await compileMarkdown("A plain paragraph.\n", {
    documentId: "motion",
    profile: "us-district-conventional",
    metadata,
  });
  const inspected = await inspectDocx(compiled.bytes);
  assert.equal(inspected.inspectOnly, true);
  assert.equal(inspected.fidelity.overall, "normalized");
  assert.equal(inspected.recognized.blocks.length, 1);
});

test("compilation embeds images and emits exhibit attachment bundles", async () => {
  const seal = Uint8Array.from(
    Buffer.from(
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4z8DwHwAFgAI/ScL77QAAAABJRU5ErkJggg==",
      "base64",
    ),
  );
  const compiled = await compileMarkdown(
    [
      '::image{source="seal.png" alt="Court seal" widthTwips="1440" heightTwips="1440"}',
      "",
      ':::exhibit{id="exhibit-a" label="Exhibit A" source="record.pdf"}',
      "",
      "Attached record.",
      "",
      ":::",
    ].join("\n"),
    {
      documentId: "motion",
      profile: "us-district-conventional",
      metadata,
      assets: {
        "seal.png": { bytes: seal, mediaType: "image/png" },
        "record.pdf": {
          bytes: new TextEncoder().encode("%PDF-1.7\nrecord\n"),
          mediaType: "application/pdf",
        },
      },
    },
  );
  assert.ok(compiled.attachments);
  assert.deepEqual(compiled.attachments.manifest.entries, [
    {
      name: "record.pdf",
      mediaType: "application/pdf",
      byteLength: 16,
      sha256:
        "sha256:433e0e77106688bce4557b2a90b88f514d3ab2dccdcd181c608db95565332af3",
      payloadPath: "files/record.pdf",
    },
  ]);
  assert.deepEqual(
    compiled.artifact.attachments?.manifest,
    compiled.attachments.manifest,
  );
  const parts = await readDocxParts(compiled.bytes);
  assert.ok([...parts.keys()].some((path) => path.startsWith("word/media/")));
  assert.match(
    new TextDecoder().decode(parts.get("word/document.xml")),
    /Exhibit A/,
  );
});

test("semantic DOCX import requires and verifies declared attachment bundles", async () => {
  const record = new TextEncoder().encode("%PDF-1.7\nrecord\n");
  const compiled = await compileMarkdown(
    ':::exhibit{id="exhibit-a" label="Exhibit A" source="record.pdf"}\n\nAttached record.\n\n:::',
    {
      documentId: "motion",
      profile: "us-district-conventional",
      metadata,
      assets: { "record.pdf": { bytes: record, mediaType: "application/pdf" } },
    },
  );
  assert.ok(compiled.attachments);
  const withoutBundle = await inspectDocx(compiled.bytes);
  assert.equal(withoutBundle.fidelity.overall, "unsupported");
  const inspected = await inspectDocx(compiled.bytes, {
    attachments: {
      manifest: compiled.attachments.manifest,
      files: compiled.attachments.files,
    },
  });
  assert.equal(inspected.fidelity.overall, "normalized");
  assert.equal(
    inspected.fidelity.items.filter((item) => item.status === "externalized")
      .length,
    1,
  );
  assert.equal(
    inspected.recognized.assets["record.pdf"]?.sha256,
    compiled.attachments.manifest.entries[0]?.sha256,
  );
});

test("attachment bundles enforce file and path budgets", async () => {
  const directory = await mkdtemp(join(tmpdir(), "agent-docx-attachments-"));
  try {
    const oversized = new Uint8Array(26 * 1024 * 1024);
    const compiled = await compileMarkdown(
      ':::exhibit{id="exhibit-a" label="Exhibit A" source="big.pdf"}\n\nAttached record.\n\n:::',
      {
        documentId: "motion",
        profile: "us-district-conventional",
        metadata,
        assets: {
          "big.pdf": { bytes: oversized, mediaType: "application/pdf" },
        },
      },
    );
    assert.ok(compiled.attachments);
    await mkdir(join(directory, "files"), { recursive: true });
    await writeFile(
      join(directory, "manifest.json"),
      JSON.stringify(compiled.attachments.manifest),
    );
    await writeFile(join(directory, "files/big.pdf"), Buffer.from(oversized));
    await assert.rejects(
      inspectDocxMaterial(compiled.bytes, {
        attachments: { directory },
      }),
      (error) =>
        error instanceof AgentDocxError && error.code === "DOCX_TOO_LARGE",
    );

    const parts = await readDocxParts(compiled.bytes);
    const manifestXml = new TextDecoder().decode(
      parts.get("customXml/itemAgentDocx.xml"),
    );
    const tampered = manifestXml.replace('"files/big.pdf"', '"files//big.pdf"');
    assert.notEqual(tampered, manifestXml);
    const tamperedParts = new Map(parts);
    tamperedParts.set(
      "customXml/itemAgentDocx.xml",
      new TextEncoder().encode(tampered),
    );
    const tamperedBytes = repackDocxParts(tamperedParts);
    await assert.rejects(
      inspectDocxMaterial(tamperedBytes, {
        attachments: { directory },
      }),
      (error) =>
        error instanceof AgentDocxError &&
        error.code === "DOCX_IMPORT_UNSUPPORTED" &&
        /version-1 shape/.test(error.message),
    );
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});

test("compilation emits native legal styles and document chrome", async () => {
  const compiled = await compileMarkdown(
    "# Argument\n\nThe requested relief follows.\n",
    {
      documentId: "motion",
      profile: "us-district-conventional",
      metadata,
      chrome: {
        headers: { default: "{{caseName}} — {{documentTitle}}" },
        footers: { default: "{{docketNumber}} — Page {{page}} of {{pages}}" },
        pageNumber: {
          story: "footer",
          alignment: "center",
          format: "decimal",
          start: 1,
        },
        lineNumbers: {
          countBy: 1,
          start: 1,
          distanceTwips: 360,
          restart: "continuous",
        },
      },
    },
  );
  const parts = await readDocxParts(compiled.bytes);
  const xml = (path) => new TextDecoder().decode(parts.get(path));
  assert.match(xml("word/styles.xml"), /AgentDocxBody/);
  assert.match(xml("word/styles.xml"), /AgentDocxHeading1/);
  assert.match(xml("word/header1.xml"), /Example v\. Example/);
  assert.match(xml("word/footer1.xml"), /Page/);
  assert.match(xml("word/document.xml"), /w:lnNumType/);
});

test("document chrome constrains both deterministic and DOCX body bounds", async () => {
  const compiled = await compileMarkdown("A plain paragraph.\n", {
    documentId: "motion",
    profile: "us-district-conventional",
    metadata,
    chrome: {
      headers: { default: "one\ntwo\nthree\nfour\nfive" },
    },
  });
  assert.ok(compiled.measurement.deterministic.lastPage.usableTwips < 12960);
  const parts = await readDocxParts(compiled.bytes);
  const documentXml = new TextDecoder().decode(parts.get("word/document.xml"));
  const topMargin = Number(documentXml.match(/w:top="(\d+)"/)?.[1]);
  assert.ok(topMargin > 1440);
});

test("clean compilation carries a deterministic semantic manifest", async () => {
  const source = "# Motion\n\nThe requested relief follows.\n";
  const options = {
    documentId: "motion",
    profile: "us-district-conventional",
    metadata,
  };
  const first = await compileMarkdown(source, options);
  const second = await compileMarkdown(source, options);
  assert.deepEqual(first.bytes, second.bytes);
  const parts = await readDocxParts(first.bytes);
  const manifest = new TextDecoder().decode(
    parts.get("customXml/itemAgentDocx.xml"),
  );
  assert.match(manifest, /semantic-manifest\/v1/);
  assert.match(manifest, /"documentId":"motion"/);
  assert.match(
    new TextDecoder().decode(parts.get("word/_rels/document.xml.rels")),
    /agent-docx\.dev\/relationships\/semantic-manifest/,
  );
});

test("semantic DOCX import restores source and validates emitted bookmarks", async () => {
  const compiled = await compileMarkdown(
    "# Motion\n\nThe requested relief follows.\n",
    {
      documentId: "motion",
      profile: "us-district-conventional",
      metadata,
    },
  );
  const inspected = await inspectDocxMaterial(compiled.bytes);
  assert.match(inspected.source, /agent-docx:block/);
  assert.deepEqual(
    inspected.result.recognized.blocks.map((block) => block.id),
    compiled.blocks.map((block) => block.id),
  );
  assert.ok(
    inspected.result.fidelity.items.some(
      (item) => item.ooxmlKind === "agent-docx:semantic-manifest",
    ),
  );
});
