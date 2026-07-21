import assert from "node:assert/strict";
import test from "node:test";
import yauzl from "yauzl";
import { generateDocx } from "../dist/docx/generate.js";
import { normalizeMarkdown } from "../dist/markdown.js";
import { builtInProfiles } from "../dist/index.js";

function zipEntries(bytes, wanted) {
  return new Promise((resolve, reject) => {
    yauzl.fromBuffer(
      Buffer.from(bytes),
      { lazyEntries: true },
      (error, zip) => {
        if (error) return reject(error);
        const entries = {};
        zip.once("error", reject);
        zip.once("end", () => resolve(entries));
        zip.on("entry", (entry) => {
          if (!wanted.includes(entry.fileName)) {
            zip.readEntry();
            return;
          }
          zip.openReadStream(entry, (streamError, stream) => {
            if (streamError) return reject(streamError);
            const chunks = [];
            stream.on("data", (chunk) => chunks.push(chunk));
            stream.once("error", reject);
            stream.once("end", () => {
              entries[entry.fileName] = Buffer.concat(chunks).toString("utf8");
              zip.readEntry();
            });
          });
        });
        zip.readEntry();
      },
    );
  });
}

function paragraphProperties(xml, text) {
  const textIndex = xml.indexOf(text);
  assert.notEqual(textIndex, -1, `text ${text} exists`);
  const starts = [...xml.slice(0, textIndex).matchAll(/<w:p(?: [^>]*)?>/g)];
  const start = starts.at(-1)?.index;
  assert.notEqual(start, undefined, `paragraph containing ${text} exists`);
  const end = xml.indexOf("</w:p>", textIndex);
  assert.notEqual(end, -1, `paragraph containing ${text} closes`);
  const paragraph = xml.slice(start, end + "</w:p>".length);
  const properties = paragraph.match(/<w:pPr>[\s\S]*?<\/w:pPr>/)?.[0];
  assert.ok(properties, `paragraph containing ${text} has properties`);
  return properties;
}

function assertProperties(properties, expected) {
  assert.match(properties, /<w:keepNext\/>/);
  assert.match(properties, /<w:keepLines\/>/);
  assert.match(properties, /<w:widowControl\/>/);
  for (const [name, value] of Object.entries(expected)) {
    assert.match(properties, new RegExp(`w:${name}="${value}"`));
  }
}

test("main and footnote paragraphs share native pagination properties", async () => {
  const profile = structuredClone(builtInProfiles["us-district-conventional"]);
  Object.assign(profile.body, {
    beforeTwips: 11,
    afterTwips: 22,
    leftIndentTwips: 33,
    rightIndentTwips: 44,
    firstLineIndentTwips: 55,
    hangingIndentTwips: 66,
    keepWithNext: true,
    keepLines: true,
    lineSpacing: { rule: "exact", twips: 480 },
  });
  Object.assign(profile.footnote, {
    beforeTwips: 77,
    afterTwips: 88,
    leftIndentTwips: 99,
    rightIndentTwips: 111,
    firstLineIndentTwips: 122,
    hangingIndentTwips: 133,
    keepWithNext: true,
    keepLines: true,
    lineSpacing: { rule: "atLeast", twips: 360 },
  });

  const bytes = await generateDocx(
    normalizeMarkdown("Main.[^1]\n\n[^1]: Footnote."),
    profile,
  );
  const entries = await zipEntries(bytes, [
    "word/document.xml",
    "word/footnotes.xml",
  ]);
  const documentXml = entries["word/document.xml"];
  const footnotesXml = entries["word/footnotes.xml"];
  assert.ok(documentXml);
  assert.ok(footnotesXml);
  assert.match(documentXml, /<w:footnoteReference w:id="1"\/>/);
  assert.match(footnotesXml, /<w:footnote w:id="1">/);

  assertProperties(paragraphProperties(documentXml, "Main."), {
    before: 11,
    after: 22,
    line: 480,
    lineRule: "exact",
    left: 33,
    right: 44,
    firstLine: 55,
    hanging: 66,
  });
  assertProperties(paragraphProperties(footnotesXml, "Footnote."), {
    before: 77,
    after: 88,
    line: 360,
    lineRule: "atLeast",
    left: 99,
    right: 111,
    firstLine: 122,
    hanging: 133,
  });
});

test("DOCX paragraphs can disable native widow and orphan control", async () => {
  const profile = structuredClone(builtInProfiles["us-district-conventional"]);
  profile.pagination.widowOrphanControl = false;

  const bytes = await generateDocx(
    normalizeMarkdown("Main.[^1]\n\n[^1]: Footnote."),
    profile,
  );
  const entries = await zipEntries(bytes, [
    "word/document.xml",
    "word/footnotes.xml",
  ]);

  for (const [xml, text] of [
    [entries["word/document.xml"], "Main."],
    [entries["word/footnotes.xml"], "Footnote."],
  ]) {
    assert.doesNotMatch(
      paragraphProperties(xml, text),
      /<w:widowControl(?:\/>| w:val="true"\/>)/,
    );
  }
});
