import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { Document, Packer, Paragraph } from "docx";
import {
  builtInProfiles,
  estimateMarkdown,
  inspectDocxTemplate,
  measureMarkdown,
  AgentDocxError,
} from "../dist/index.js";

test("root API and immutable profiles", () => {
  assert.deepEqual(Object.keys(builtInProfiles), [
    "us-district-conventional",
    "frap-32",
    "cand-civil",
  ]);
  assert.equal(Object.isFrozen(builtInProfiles), true);
  assert.equal(typeof estimateMarkdown, "function");
  assert.equal(typeof measureMarkdown, "function");
  assert.equal(typeof inspectDocxTemplate, "function");
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
    () => estimateMarkdown("| A |\n| - |\n| x |"),
    (error) =>
      error instanceof AgentDocxError &&
      error.code === "UNSUPPORTED_MARKDOWN",
  );
  await assert.rejects(
    () => estimateMarkdown("```js\nalert(1)\n```"),
    (error) =>
      error instanceof AgentDocxError &&
      error.code === "UNSUPPORTED_MARKDOWN",
  );
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
    "Shortening or rephrasing this paragraph may remove its final wrapped line.",
  );
  assert.equal(result.budget.limitPages, 1);
});

test("block style, indentation, and line-cap exclusions affect deterministic layout", async () => {
  const bold = await estimateMarkdown("# " + "Legal filing ".repeat(3), {
    layout: {
      page: { widthTwips: 2800, marginsTwips: { left: 500, right: 500 } },
      headings: { 1: { bold: true } },
    },
  });
  const regular = await estimateMarkdown("# " + "Legal filing ".repeat(3), {
    layout: {
      page: { widthTwips: 2800, marginsTwips: { left: 500, right: 500 } },
      headings: { 1: { bold: false } },
    },
  });
  assert.notEqual(bold.totalVisualLines, regular.totalVisualLines);

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
  assert.equal(continuing.paragraphs[1].startPage, 3);
  assert.equal(relaxedWarnings(continuing).length, 1);
});
