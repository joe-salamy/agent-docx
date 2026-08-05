import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { Document, Packer, Paragraph } from "docx";
import {
  builtInProfiles,
  compileMarkdown,
  estimateMarkdown,
  inspectDocxTemplate,
  lowerLegalDocument,
  measureMarkdown,
  parseLegalMarkdown,
  AgentDocxError,
} from "../dist/index.js";
import { measureNormalizedDocument } from "../dist/renderers/index.js";
import { normalizeMarkdown } from "../dist/markdown.js";
import { metadata } from "./helpers.js";

test("root API and immutable profiles", () => {
  assert.deepEqual(Object.keys(builtInProfiles), [
    "us-district-conventional",
    "frap-32",
    "cand-civil",
  ]);
  assert.equal(Object.isFrozen(builtInProfiles), true);
  assert.equal(
    builtInProfiles["us-district-conventional"].pagination.widowOrphanControl,
    true,
  );
  assert.equal(typeof estimateMarkdown, "function");
  assert.equal(typeof measureMarkdown, "function");
  assert.equal(typeof inspectDocxTemplate, "function");
});

test("generated DOCX bytes are opt-in and publicly inspectable", async () => {
  const regular = await measureMarkdown("# Generated");
  assert.equal("generatedDocx" in regular, false);
  const included = await measureMarkdown("# Generated", {
    includeGeneratedDocx: true,
  });
  assert.ok(included.generatedDocx instanceof Uint8Array);
  assert.equal(
    Buffer.from(included.generatedDocx).subarray(0, 2).toString(),
    "PK",
  );
  await inspectDocxTemplate(included.generatedDocx);
});

test("image blocks reserve their extent in deterministic pagination", async () => {
  const seal = Uint8Array.from(
    Buffer.from(
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4z8DwHwAFgAI/ScL77QAAAABJRU5ErkJggg==",
      "base64",
    ),
  );
  const parseOptions = {
    documentId: "motion",
    profile: "us-district-conventional",
    metadata: {
      court: "",
      jurisdiction: "",
      caseName: "",
      docketNumber: "",
      documentTitle: "",
      parties: [],
      counsel: [],
      certificates: [],
    },
    assets: { "seal.png": { bytes: seal, mediaType: "image/png" } },
  };
  const measure = (markdown) =>
    measureNormalizedDocument(
      lowerLegalDocument(parseLegalMarkdown(markdown, parseOptions).document),
      { renderer: "deterministic" },
    );
  const one = await measure(
    'Intro.\n\n::image{source="seal.png" alt="Seal" widthTwips="2000" heightTwips="6500"}\n',
  );
  assert.equal(one.deterministic.pageCount, 1);
  assert.ok(
    one.deterministic.lastPage.usedTwips >= 6500,
    `last page usedTwips ${one.deterministic.lastPage.usedTwips} must include the image reservation`,
  );
  const two = await measure(
    [
      '::image{source="seal.png" alt="Seal A" widthTwips="2000" heightTwips="6500"}',
      "",
      '::image{source="seal.png" alt="Seal B" widthTwips="2000" heightTwips="6500"}',
    ].join("\n"),
  );
  assert.equal(two.deterministic.pageCount, 2);
});

test("default metric fonts match the committed manifest", async () => {
  const manifest = JSON.parse(
    await readFile(
      new URL(
        "../assets/fonts/liberation-serif-2.1.5/manifest.json",
        import.meta.url,
      ),
      "utf8",
    ),
  );
  const result = await estimateMarkdown("Manifest check.");
  assert.deepEqual(
    result.profile.metricFonts.map(({ role, metricsFamily, sha256 }) => ({
      role,
      metricsFamily,
      sha256,
    })),
    manifest.fonts.map(({ sha256 }, index) => ({
      role: ["regular", "bold", "italic", "boldItalic"][index],
      metricsFamily: "Liberation Serif",
      sha256,
    })),
  );
});

test("empty Markdown has no deterministic physical page", async () => {
  const result = await estimateMarkdown(" \n\t");
  assert.equal(result.pageCount, 0);
  assert.equal(result.equivalentPages, 0);
  assert.equal(result.lastPage, null);
});

test("exact 27-line boundary and 28th-line overflow", async () => {
  const markdown27 = await readFile(
    new URL("./fixtures/27-hard-lines.md", import.meta.url),
    "utf8",
  );
  const markdown28 = await readFile(
    new URL("./fixtures/28-hard-lines.md", import.meta.url),
    "utf8",
  );
  const options = {
    layout: {
      body: {
        lineSpacing: { rule: "exact", twips: 480 },
        beforeTwips: 0,
        afterTwips: 0,
        keepLines: false,
      },
      pagination: { widowLines: 1, orphanLines: 1 },
    },
  };
  const one = await estimateMarkdown(markdown27, options);
  const two = await estimateMarkdown(markdown28, options);
  assert.equal(one.pageCount, 1);
  assert.equal(two.pageCount, 2);
  assert.equal(two.equivalentPages, 1 + 480 / 12960);
  assert.equal(two.lastPage.bodyLineEquivalentsUsed, 1);
  assert.equal(two.lastPage.bodyLineCapacity, 27);
});

test("explicit page break abandons remaining page", async () => {
  const result = await estimateMarkdown(
    "First.\n\n<!-- pagebreak -->\n\nSecond.",
  );
  assert.equal(result.pageCount, 2);
  assert.ok(result.equivalentPages > 1 && result.equivalentPages < 2);
});

test("unsupported Markdown rejects with source-aware code", async () => {
  await assert.rejects(
    () => estimateMarkdown("---\ntitle: filing\n---\n\nBody."),
    (error) =>
      error instanceof AgentDocxError &&
      error.code === "UNSUPPORTED_MARKDOWN" &&
      error.details.position.start.line === 1,
  );
  await assert.rejects(
    () => estimateMarkdown("```js\nalert(1)\n```"),
    (error) =>
      error instanceof AgentDocxError &&
      error.code === "UNSUPPORTED_MARKDOWN",
  );
  for (const [markdown, line] of [
    ["| A |\n| - |\n| ![image](file.png) |", 3],
    ["| A |\n| - |\n| Ref.[^1] |\n\n[^1]: Note.", 3],
    ["| A |\n| - |\n| <span>HTML</span> |", 3],
    ["[^1]:\n    - list child\n\nBody.[^1]", 2],
    ["Inline $$x + y$$ math.", 1],
    ["$$\nx + y\n$$", 1],
  ]) {
    await assert.rejects(
      () => estimateMarkdown(markdown),
      (error) =>
        error instanceof AgentDocxError &&
        error.code === "UNSUPPORTED_MARKDOWN" &&
        error.details.position.start.line === line,
    );
  }
});

test("font shaping distinguishes narrow and wide glyphs", async () => {
  const narrow = await estimateMarkdown("iiii", { paragraphDiagnostics: true });
  const wide = await estimateMarkdown("WWWW", { paragraphDiagnostics: true });
  assert.ok(
    wide.paragraphs[0].lastLineUsedTwips >
      narrow.paragraphs[0].lastLineUsedTwips,
  );
});

test("CAND independent 28-counted-line ceiling", async () => {
  const hard = (n) =>
    Array.from({ length: n }, (_, i) => `x${i < n - 1 ? "  " : ""}`).join("\n");
  const layout = {
    body: { lineSpacing: { rule: "exact", twips: 400 }, keepLines: false },
    pagination: { widowLines: 1, orphanLines: 1 },
  };
  assert.equal(
    (await estimateMarkdown(hard(28), { profile: "cand-civil", layout }))
      .pageCount,
    1,
  );
  assert.equal(
    (await estimateMarkdown(hard(29), { profile: "cand-civil", layout }))
      .pageCount,
    2,
  );
});

test("trim diagnostics are deterministic and advisory", async () => {
  const result = await estimateMarkdown("Word ".repeat(250), {
    paragraphDiagnostics: true,
    trim: { maxLastLineRatio: 1, maxCandidates: 10 },
    pageLimit: 1,
  });
  assert.ok(result.paragraphs.length === 1);
  const p = result.paragraphs[0];
  assert.equal(p.lastLineRatio, p.lastLineUsedTwips / p.lastLineAvailableTwips);
  assert.equal(
    result.trimOpportunities[0].message,
    "This block may lose one wrapped line after removing or rephrasing approximately the reported width; verify by re-running pagination.",
  );
  assert.ok(p.lastLineText.length > 0);
  assert.equal(
    p.lastLineUnusedTwips,
    p.lastLineAvailableTwips - p.lastLineUsedTwips,
  );
  assert.equal(p.lastLineOverflow, false);
  assert.equal(p.oneLineReduction.confidence, "heuristic");
  assert.equal(result.budget.limitPages, 1);
});

test("paragraph tails retain exact and node source provenance", async () => {
  const markdown = "Alpha beta gamma delta epsilon.";
  const exact = await estimateMarkdown(markdown, {
    paragraphDiagnostics: true,
    layout: {
      page: { widthTwips: 3600, marginsTwips: { left: 1440, right: 1440 } },
    },
  });
  const tail = exact.paragraphs[0];
  assert.ok(tail.visualLines > 1);
  assert.equal(
    markdown.slice(
      tail.lastLineSourceRanges[0].position.start.offset,
      tail.lastLineSourceRanges[0].position.end.offset,
    ),
    tail.lastLineText,
  );
  assert.ok(
    tail.lastLineSourceRanges.every(({ precision }) => precision === "exact"),
  );
  assert.deepEqual(tail.lastLineTextRange, {
    start: markdown.indexOf(tail.lastLineText),
    end: markdown.indexOf(tail.lastLineText) + tail.lastLineText.length,
  });

  const decoded = await estimateMarkdown("Alpha &amp; omega.", {
    paragraphDiagnostics: true,
  });
  assert.equal(decoded.paragraphs[0].oneLineReduction, null);
  assert.ok(
    decoded.paragraphs[0].lastLineSourceRanges.some(
      ({ precision }) => precision === "node",
    ),
  );

  const normalizedBreak = normalizeMarkdown("First.  \nSecond.").blocks[0];
  assert.ok(
    normalizedBreak.sourceSegments.some(
      ({ precision }) => precision === "node",
    ),
  );
  const hardBreak = await estimateMarkdown("First.  \nSecond.", {
    paragraphDiagnostics: true,
  });
  assert.equal(hardBreak.paragraphs[0].lastLineText, "Second.");
  assert.equal(hardBreak.paragraphs[0].oneLineReduction, null);
});

test("overflow and trim ranking expose signed deterministic width facts", async () => {
  const overflow = await estimateMarkdown("W".repeat(80), {
    paragraphDiagnostics: true,
    layout: {
      page: { widthTwips: 3100, marginsTwips: { left: 1440, right: 1440 } },
    },
  });
  assert.equal(overflow.paragraphs[0].lastLineOverflow, true);
  assert.ok(overflow.paragraphs[0].lastLineUnusedTwips < 0);

  const ranked = await estimateMarkdown(
    [
      "Alpha beta gamma delta epsilon zeta eta theta.",
      "One two three four five six seven eight nine ten eleven.",
      "Legal legal legal legal legal legal legal legal legal.",
    ].join("\n\n"),
    {
      trim: { maxCandidates: 10, maxLastLineRatio: 1 },
      layout: {
        page: {
          widthTwips: 4200,
          marginsTwips: { left: 1440, right: 1440 },
        },
      },
    },
  );
  const widths = ranked.trimOpportunities.map(
    ({ oneLineReduction }) => oneLineReduction.estimatedRemovalTwips,
  );
  assert.deepEqual(
    widths,
    [...widths].sort((a, b) => a - b),
  );
});

test("tables and thematic breaks use structural layout without paragraph diagnostics", async () => {
  const table = "| A | B |\n| :- | -: |\n| *x* | `y` |";
  const exactFloor = await estimateMarkdown(table, {
    paragraphDiagnostics: true,
    layout: {
      page: {
        widthTwips: 3210,
        marginsTwips: { left: 1440, right: 1440 },
      },
    },
  });
  assert.equal(exactFloor.paragraphs.length, 0);
  assert.equal(exactFloor.totalVisualLines, 2);
  await assert.rejects(
    () =>
      estimateMarkdown(table, {
        layout: {
          page: {
            widthTwips: 3209,
            marginsTwips: { left: 1440, right: 1440 },
          },
        },
      }),
    (error) =>
      error instanceof AgentDocxError &&
      error.code === "INVALID_LAYOUT" &&
      error.details.position.start.line === 1,
  );

  const repeated = await estimateMarkdown(
    "| Header |\n| --- |\n| Row one |\n| Row two |\n| Row three |",
    {
      layout: {
        page: {
          heightTwips: 1440,
          marginsTwips: { top: 0, bottom: 0 },
          headerTwips: 0,
          footerTwips: 0,
        },
      },
    },
  );
  assert.equal(repeated.pageCount, 3);
  assert.deepEqual(repeated.visualLinesByPage, [2, 2, 2]);
  assert.equal(
    repeated.warnings.some(
      ({ code }) => code === "TABLE_HEADER_REPEAT_CONSTRAINT_RELAXED",
    ),
    false,
  );

  const atomic = await estimateMarkdown(
    "| Header |\n| --- |\n| " + "wrapped cell ".repeat(3) + " |",
    {
      layout: {
        page: {
          widthTwips: 5000,
          heightTwips: 2160,
          marginsTwips: { top: 0, right: 1440, bottom: 0, left: 1440 },
          headerTwips: 0,
          footerTwips: 0,
        },
      },
    },
  );
  assert.equal(atomic.pageCount, 2);
  assert.equal(atomic.visualLinesByPage[0], 1);
  assert.ok(atomic.visualLinesByPage[1] > 1);
  assert.equal(
    atomic.warnings.some(
      ({ code }) => code === "TABLE_ROW_SPLIT_CONSTRAINT_RELAXED",
    ),
    false,
  );

  const keptRule = await estimateMarkdown("Before.\n\n---\n\nAfter.", {
    layout: {
      page: {
        heightTwips: 960,
        marginsTwips: { top: 0, bottom: 0 },
        headerTwips: 0,
        footerTwips: 0,
      },
    },
  });
  assert.equal(keptRule.pageCount, 2);
  assert.ok(keptRule.lastPage.usedTwips > 480);

  const plain = await estimateMarkdown("Before.\n\nAfter.");
  const ruled = await estimateMarkdown("Before.\n\n---\n\nAfter.", {
    paragraphDiagnostics: true,
  });
  assert.equal(ruled.paragraphs.length, 2);
  assert.equal(ruled.totalVisualLines, plain.totalVisualLines);
  assert.ok(ruled.equivalentPages > plain.equivalentPages);
});

test("multi-paragraph footnotes preserve child order and child spacing", async () => {
  const result = await estimateMarkdown(
    "Body.[^1]\n\n[^1]: First child.\n    \n    Second child.",
    { layout: { footnote: { beforeTwips: 40, afterTwips: 60 } } },
  );
  assert.equal(result.pageCount, 1);
  assert.equal(result.totalVisualLines, 3);
  assert.ok(result.lastPage.usedTwips > 0);
  assert.equal(
    result.warnings.filter(
      ({ code }) => code === "FOOTNOTE_SPLIT_CONSTRAINT_RELAXED",
    ).length,
    0,
  );
});

test("block style, indentation, and line-cap exclusions affect deterministic layout", async () => {
  const bold = await estimateMarkdown("# " + "Legal filing ".repeat(3), {
    paragraphDiagnostics: true,
    layout: {
      page: { widthTwips: 2800, marginsTwips: { left: 500, right: 500 } },
      headings: { 1: { bold: true } },
    },
  });
  const regular = await estimateMarkdown("# " + "Legal filing ".repeat(3), {
    paragraphDiagnostics: true,
    layout: {
      page: { widthTwips: 2800, marginsTwips: { left: 500, right: 500 } },
      headings: { 1: { bold: false } },
    },
  });
  assert.notEqual(
    bold.paragraphs[0].lastLineUsedTwips,
    regular.paragraphs[0].lastLineUsedTwips,
  );

  const plain = await estimateMarkdown("Short line.", {
    paragraphDiagnostics: true,
  });
  const indented = await estimateMarkdown("Short line.", {
    paragraphDiagnostics: true,
    layout: { body: { firstLineIndentTwips: 720 } },
  });
  assert.equal(
    indented.paragraphs[0].lastLineAvailableTwips,
    plain.paragraphs[0].lastLineAvailableTwips - 720,
  );

  const markdown = "> First line.  \n> Second line.";
  const pagination = {
    maxCountedLinesPerPage: 1,
    widowLines: 1,
    orphanLines: 1,
  };
  assert.equal(
    (
      await estimateMarkdown(markdown, {
        layout: { pagination: { ...pagination, lineCapExclusions: [] } },
      })
    ).pageCount,
    2,
  );
  assert.equal(
    (
      await estimateMarkdown(markdown, {
        layout: {
          pagination: { ...pagination, lineCapExclusions: ["blockquote"] },
        },
      })
    ).pageCount,
    1,
  );
});

test("generated DOCX inspection imports section geometry", async () => {
  const doc = new Document({
    sections: [
      {
        properties: {
          page: {
            size: { width: 11907, height: 16839 },
            margin: {
              top: 1000,
              right: 1100,
              bottom: 1200,
              left: 1300,
              header: 500,
              footer: 600,
              gutter: 0,
            },
          },
        },
        children: [new Paragraph("Template")],
      },
    ],
  });
  const bytes = await Packer.toBuffer(doc);
  const inspected = await inspectDocxTemplate(bytes);
  assert.equal(inspected.package.macroEnabled, false);
  assert.equal(inspected.sections.at(-1).page.widthTwips, 11907);
  assert.equal(inspected.imported.page.marginsTwips.left, 1300);
});

test("template inspection resolves paragraph style inheritance into pagination", async () => {
  const template = await inspectDocxTemplate(
    await readFile("test/fixtures/docx/theme-inheritance.docx"),
  );
  assert.equal(template.styles.headings["1"].resolved.fontSizeTwips, 320);
  assert.equal(template.styles.headings["1"].provenance[""], "template");
  const measured = await estimateMarkdown("# Heading\n", { template });
  assert.equal(measured.profile.headings["1"].fontSizeTwips, 320);
  assert.equal(measured.profile.requestedFontFamily, "Times New Roman");
});

test("template inspection reports numbering, chrome fields, and captions", async () => {
  const compiled = await compileMarkdown("::caption\n\n- One\n- Two\n", {
    documentId: "motion",
    profile: "us-district-conventional",
    metadata,
    chrome: {
      headers: { default: "{{caseName}} {{page}}" },
      footers: { default: "{{documentTitle}} {{pages}}" },
    },
  });
  const inspected = await inspectDocxTemplate(compiled.bytes);
  assert.ok(inspected.numbering.abstractNumbers.length > 0);
  assert.ok(inspected.numbering.instances.length > 0);
  assert.ok(
    inspected.headerFooters.some(
      (entry) =>
        entry.kind === "header" &&
        entry.text.includes("Example v. Example") &&
        entry.fields.some((field) => field.kind === "PAGE"),
    ),
  );
  assert.ok(inspected.fields.some((field) => field.kind === "NUMPAGES"));
  assert.ok(inspected.captions.some((caption) => caption.styleId === "AgentDocxCaption"));
  assert.equal(inspected.styles.list.resolved.leftIndentTwips, 720);
  assert.deepEqual(inspected.unsupportedParts, []);
});

test("page fields converge body bounds across a page-count digit boundary", async () => {
  const result = await estimateMarkdown(
    Array.from({ length: 20 }, () => "Line.").join("\n\n"),
    {
      layout: tinyLayout(2400),
      chrome: {
        headers: { default: `${"x".repeat(101)}{{pages}}` },
      },
    },
  );
  assert.equal(result.pageCount, 20);
  assert.equal(result.lastPage.usableTwips, 720);
});

function tinyLayout(heightTwips = 960, overrides = {}) {
  return {
    page: {
      heightTwips,
      marginsTwips: { top: 0, bottom: 0, left: 0, right: 0 },
    },
    body: {
      lineSpacing: { rule: "exact", twips: 480 },
      beforeTwips: 0,
      afterTwips: 0,
      keepLines: false,
    },
    headings: {
      1: {
        lineSpacing: { rule: "exact", twips: 480 },
        beforeTwips: 0,
        afterTwips: 0,
        keepLines: false,
        keepWithNext: true,
      },
    },
    footnote: {
      lineSpacing: { rule: "exact", twips: 480 },
      beforeTwips: 0,
      afterTwips: 0,
      keepLines: false,
      ...overrides.footnote,
    },
    pagination: {
      widowLines: 1,
      orphanLines: 1,
      widowOrphanControl: true,
      maxCountedLinesPerPage: null,
      lineCapExclusions: ["footnote"],
      ...overrides.pagination,
    },
  };
}

const relaxedWarnings = (result) =>
  result.warnings.filter(
    ({ code }) => code === "FOOTNOTE_SPLIT_CONSTRAINT_RELAXED",
  );

test("feasible keep-with-next chains move without blank pages", async () => {
  const options = { paragraphDiagnostics: true, layout: tinyLayout() };
  const feasible = await estimateMarkdown(
    "Filler.\n\n# Heading\n\nBody.",
    options,
  );
  assert.equal(feasible.pageCount, 2);
  assert.equal(feasible.lastPage.visualLines, 2);
  assert.equal(feasible.paragraphs[1].startPage, 2);

  const oversized = await estimateMarkdown("# One\n\n# Two\n\nBody.", options);
  assert.equal(oversized.pageCount, 2);
  assert.equal(oversized.totalVisualLines, 3);
  assert.equal(oversized.lastPage.visualLines, 2);

  const beforeBreak = await estimateMarkdown(
    "# Heading\n\n<!-- pagebreak -->\n\nBody.",
    options,
  );
  assert.equal(beforeBreak.pageCount, 2);
  assert.equal(beforeBreak.lastPage.visualLines, 1);

  const atEnd = await estimateMarkdown("Filler.\n\n# Heading", options);
  assert.equal(atEnd.pageCount, 1);
  assert.equal(atEnd.lastPage.visualLines, 2);
});

test("footnotes trigger on the exact wrapped line and reserve the page bottom", async () => {
  const result = await estimateMarkdown(
    "First.  \nSecond.[^1]\n\n[^1]: Note.",
    {
      paragraphDiagnostics: true,
      layout: tinyLayout(),
    },
  );
  assert.equal(result.pageCount, 2);
  assert.equal(result.paragraphs[0].startPage, 1);
  assert.equal(result.paragraphs[0].endPage, 2);
  assert.equal(result.lastPage.visualLines, 2);
});

test("footnote definitions are first-use, transitive, and never ordinary body flow", async () => {
  const layout = tinyLayout(1440);
  const repeated = await estimateMarkdown("Body.[^1][^1]\n\n[^1]: Note.", {
    layout,
  });
  assert.equal(repeated.totalVisualLines, 2);

  const multiple = await estimateMarkdown(
    "Body.[^1][^2]\n\n[^1]: One.\n\n[^2]: Two.",
    { layout },
  );
  assert.equal(multiple.totalVisualLines, 3);

  const nested = await estimateMarkdown(
    "Body.[^a]\n\n[^a]: A.[^b]\n\n[^b]: B.",
    { layout },
  );
  assert.equal(nested.totalVisualLines, 3);

  const cyclic = await estimateMarkdown(
    "Body.[^a]\n\n[^a]: A.[^b]\n\n[^b]: B.[^a]",
    { layout },
  );
  assert.equal(cyclic.totalVisualLines, 3);

  const unreferenced = await estimateMarkdown("Body.\n\n[^unused]: Hidden.", {
    layout,
  });
  assert.equal(unreferenced.totalVisualLines, 1);

  await assert.rejects(
    () => estimateMarkdown("Body.[^a]\n\n[^a]: Nested.[^missing]", { layout }),
    (error) =>
      error instanceof AgentDocxError &&
      error.code === "UNSUPPORTED_MARKDOWN" &&
      /missing/.test(error.message),
  );
});

test("footnote continuations are finite, ordered, and warn exactly once", async () => {
  const oversized = await estimateMarkdown(
    "Reference.[^1]\n\nLater.\n\n[^1]: One.  \nTwo.  \nThree.",
    { paragraphDiagnostics: true, layout: tinyLayout() },
  );
  assert.equal(oversized.pageCount, 3);
  assert.equal(oversized.paragraphs[1].startPage, 3);
  assert.equal(relaxedWarnings(oversized).length, 1);

  const zeroShare = await estimateMarkdown(
    "Reference.[^1]\n\n[^1]: One.  \nTwo.",
    {
      layout: tinyLayout(960, {
        pagination: { maxCountedLinesPerPage: 1, lineCapExclusions: [] },
      }),
    },
  );
  assert.equal(zeroShare.pageCount, 3);
  assert.equal(relaxedWarnings(zeroShare).length, 1);
});

test("footnote line caps and split constraints honor configured semantics", async () => {
  const markdown = "Body.[^1]\n\n[^1]: Note.";
  const excluded = await estimateMarkdown(markdown, {
    layout: tinyLayout(960, {
      pagination: {
        maxCountedLinesPerPage: 1,
        lineCapExclusions: ["footnote"],
      },
    }),
  });
  const counted = await estimateMarkdown(markdown, {
    layout: tinyLayout(960, {
      pagination: { maxCountedLinesPerPage: 1, lineCapExclusions: [] },
    }),
  });
  assert.equal(excluded.pageCount, 1);
  assert.equal(counted.pageCount, 2);

  const note = "Reference.[^1]\n\nLater.\n\n[^1]: One.  \nTwo.  \nThree.";
  const feasible = await estimateMarkdown(note, {
    layout: tinyLayout(1920, { pagination: { widowLines: 2, orphanLines: 1 } }),
  });
  assert.equal(feasible.pageCount, 2);
  assert.equal(relaxedWarnings(feasible).length, 0);
  const controlledSplit = await estimateMarkdown(note, {
    layout: tinyLayout(1440, {
      pagination: { widowLines: 2, orphanLines: 1 },
    }),
  });
  const disabledSplit = await estimateMarkdown(note, {
    layout: tinyLayout(1440, {
      pagination: {
        widowOrphanControl: false,
        widowLines: 2,
        orphanLines: 1,
      },
    }),
  });
  assert.equal(controlledSplit.lastPage.visualLines, 3);
  assert.equal(disabledSplit.lastPage.visualLines, 2);

  const relaxed = await estimateMarkdown(note, {
    layout: tinyLayout(960, { pagination: { widowLines: 2, orphanLines: 2 } }),
  });
  assert.equal(relaxed.pageCount, 3);
  assert.equal(relaxedWarnings(relaxed).length, 1);

  const kept = await estimateMarkdown("Reference.[^1]\n\n[^1]: One.  \nTwo.", {
    layout: tinyLayout(960, { footnote: { keepLines: true } }),
  });
  assert.equal(kept.pageCount, 2);
  assert.equal(relaxedWarnings(kept).length, 0);
});

test("widow and orphan control can be disabled", async () => {
  const markdown = "Filler.\n\nOne.  \nTwo.  \nThree.";
  const controlled = await estimateMarkdown(markdown, {
    paragraphDiagnostics: true,
    layout: tinyLayout(1440, {
      pagination: { widowLines: 2, orphanLines: 2 },
    }),
  });
  const disabled = await estimateMarkdown(markdown, {
    paragraphDiagnostics: true,
    layout: tinyLayout(1440, {
      pagination: {
        widowOrphanControl: false,
        widowLines: 2,
        orphanLines: 2,
      },
    }),
  });

  assert.equal(controlled.paragraphs[1].startPage, 2);
  assert.equal(controlled.lastPage.visualLines, 3);
  assert.equal(disabled.paragraphs[1].startPage, 1);
  assert.equal(disabled.lastPage.visualLines, 1);
  await assert.rejects(
    () =>
      estimateMarkdown(markdown, {
        layout: { pagination: { widowOrphanControl: "false" } },
      }),
    (error) =>
      error instanceof AgentDocxError &&
      error.code === "INVALID_LAYOUT" &&
      /widowOrphanControl/.test(error.message),
  );
});

test("keep-chain preflight reserves fitting and continuing notes transactionally", async () => {
  const fitting = await estimateMarkdown(
    "Filler.\n\n# Heading[^1]\n\nBody.\n\n[^1]: Note.",
    { paragraphDiagnostics: true, layout: tinyLayout(1440) },
  );
  assert.equal(fitting.pageCount, 2);
  assert.equal(fitting.lastPage.visualLines, 3);
  assert.equal(fitting.paragraphs[1].startPage, 2);
  assert.equal(relaxedWarnings(fitting).length, 0);

  const continuing = await estimateMarkdown(
    "Filler.\n\n# Heading[^1][^1]\n\nLater.\n\n[^1]: One.  \nTwo.  \nThree.  \nFour.",
    { paragraphDiagnostics: true, layout: tinyLayout(1440) },
  );
  assert.equal(continuing.pageCount, 3);
  assert.equal(continuing.totalVisualLines, 7);
  assert.equal(
    continuing.paragraphs.find(({ preview }) => preview === "Later.").startPage,
    3,
  );
  assert.equal(relaxedWarnings(continuing).length, 1);
});

test("section indexing preserves source identity and inclusive hierarchy", async () => {
  const markdown =
    "Preamble.\n\n# Parent\n\n### Duplicate\n\nBody.\n\n## Duplicate\n\n##\n";
  const regular = await estimateMarkdown(markdown);
  assert.equal("sections" in regular, false);
  const result = await estimateMarkdown(markdown, {
    sectionDiagnostics: true,
  });
  assert.deepEqual(
    result.sections.map((section) => ({
      index: section.index,
      parent: section.parentIndex,
      title: section.heading?.title ?? null,
      empty: section.empty,
      line: section.heading?.position.start.line ?? null,
    })),
    [
      { index: 0, parent: null, title: null, empty: false, line: null },
      { index: 1, parent: null, title: "Parent", empty: false, line: 3 },
      { index: 2, parent: 1, title: "Duplicate", empty: false, line: 5 },
      { index: 3, parent: 1, title: "Duplicate", empty: true, line: 9 },
      { index: 4, parent: 1, title: "", empty: true, line: 11 },
    ],
  );
  assert.equal(result.sections[0].position.start.line, 1);
  assert.equal(result.sections[0].position.end.line, 1);
  assert.equal(result.sections[1].position.end.line, 11);
  assert.ok(
    result.sections[1].visualLines > result.sections[2].visualLines,
    "inclusive parent contains its own heading and descendants",
  );

  const headingFirst = await estimateMarkdown("## Leading", {
    sectionDiagnostics: true,
  });
  assert.equal(headingFirst.sections[0].position, null);
  assert.equal(headingFirst.sections[1].parentIndex, null);
  const empty = await estimateMarkdown("", { sectionDiagnostics: true });
  assert.equal(empty.sections.length, 1);
  assert.deepEqual(
    [empty.sections[0].position, empty.sections[0].startPage],
    [null, null],
  );
});

test("section pages include lines, footnotes, breaks, and page budgets", async () => {
  const broken = await estimateMarkdown(
    "# Parent\n\nRef.[^n]\n\n## Child\n\nBody.\n\n<!-- pagebreak -->\n\n<!-- pagebreak -->\n\n[^n]: One.  \nTwo.",
    {
      sectionDiagnostics: true,
      pageLimit: 1,
      layout: tinyLayout(),
    },
  );
  const parent = broken.sections[1];
  const child = broken.sections[2];
  assert.deepEqual(
    parent.pages.map(({ page }) => page),
    [...parent.pages.map(({ page }) => page)].sort((a, b) => a - b),
  );
  assert.equal(parent.pageCount, parent.pages.length);
  assert.ok(parent.footnoteVisualLines > 0);
  assert.equal(child.footnoteVisualLines, 0);
  assert.ok(parent.visualLines >= child.visualLines);
  assert.deepEqual(
    parent.pageBudget.pagesBeyondLimit,
    parent.pages.map(({ page }) => page).filter((page) => page > 1),
  );
  assert.equal(parent.pageBudget.withinLimit, false);

  const breaks = await estimateMarkdown(
    "# Breaks\n\n<!-- pagebreak -->\n\n<!-- pagebreak -->",
    { sectionDiagnostics: true },
  );
  assert.deepEqual(
    breaks.sections[1].pages.map(({ page }) => page),
    [1, 2],
  );
  assert.equal(breaks.pageCount, 2);
});
