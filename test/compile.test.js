import assert from "node:assert/strict";
import test from "node:test";
import {
  builtInProfiles,
  compileMarkdown,
  generateDocx,
  generateRedlineDocx,
  inspectDocx,
  parseLegalMarkdown,
} from "../dist/index.js";
import { inspectDocxMaterial } from "../dist/docx/import.js";
import { readDocxParts } from "../dist/docx/package.js";

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

test("standalone compilation parses legal Markdown and emits stable block bookmarks", async () => {
  const compiled = await compileMarkdown("# Motion\n\nThe requested relief follows.\n", {
    documentId: "motion",
    profile: "us-district-conventional",
    metadata,
  });
  assert.ok(compiled.bytes.byteLength > 1000);
  assert.equal(compiled.artifact.revision, null);
  assert.equal(compiled.validation.status, "pass");
  assert.equal(compiled.blocks.length, 2);
  for (const block of compiled.blocks)
    assert.match(block.bookmark, /^adx_[0-9a-f]{32}$/);
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

test("native redline export accepts tables, lists, and footnotes", async () => {
  const parsed = parseLegalMarkdown(
    [
      "- First item",
      "- Second item[^note]",
      "",
      "| Left | Right |",
      "| --- | --- |",
      "| One | Two |",
      "",
      "[^note]: Supporting note.",
    ].join("\n"),
    {
      documentId: "motion",
      profile: "us-district-conventional",
      metadata,
    },
  );
  const generated = await generateRedlineDocx(
    parsed.document,
    parsed.document,
    { changes: [] },
    builtInProfiles["us-district-conventional"],
  );
  const parts = await readDocxParts(generated.bytes);
  const documentXml = new TextDecoder().decode(parts.get("word/document.xml"));
  assert.match(documentXml, /First item/);
  assert.match(documentXml, /Supporting note/);
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
        counsel: [{
          id: "counsel-1",
          name: "A. Counsel",
          addressLines: [],
          phone: "",
          email: "",
        }],
        certificates: [{
          id: "certificate-1",
          kind: "compliance",
          basis: "words",
          signerCounselId: "counsel-1",
        }],
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
  assert.deepEqual(compiled.attachments.manifest.entries, [{
    name: "record.pdf",
    mediaType: "application/pdf",
    byteLength: 16,
    sha256: "sha256:433e0e77106688bce4557b2a90b88f514d3ab2dccdcd181c608db95565332af3",
    payloadPath: "files/record.pdf",
  }]);
  assert.deepEqual(compiled.artifact.attachments?.manifest, compiled.attachments.manifest);
  const parts = await readDocxParts(compiled.bytes);
  assert.ok([...parts.keys()].some((path) => path.startsWith("word/media/")));
  assert.match(new TextDecoder().decode(parts.get("word/document.xml")), /Exhibit A/);
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
    inspected.fidelity.items.filter((item) => item.status === "externalized").length,
    1,
  );
  assert.equal(
    inspected.recognized.assets["record.pdf"]?.sha256,
    compiled.attachments.manifest.entries[0]?.sha256,
  );
});


test("compilation emits native legal styles and document chrome", async () => {
  const compiled = await compileMarkdown("# Argument\n\nThe requested relief follows.\n", {
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
  });
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
  const manifest = new TextDecoder().decode(parts.get("customXml/itemAgentDocx.xml"));
  assert.match(manifest, /semantic-manifest\/v1/);
  assert.match(manifest, /"documentId":"motion"/);
  assert.match(
    new TextDecoder().decode(parts.get("word/_rels/document.xml.rels")),
    /agent-docx\.dev\/relationships\/semantic-manifest/,
  );
});

test("semantic DOCX import restores source and validates emitted bookmarks", async () => {
  const compiled = await compileMarkdown("# Motion\n\nThe requested relief follows.\n", {
    documentId: "motion",
    profile: "us-district-conventional",
    metadata,
  });
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