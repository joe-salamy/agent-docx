import { readFile, writeFile } from "node:fs/promises";
import { createHash } from "node:crypto";
import { measureMarkdown, AgentDocxError } from "../dist/index.js";
const argIndex = process.argv.indexOf("--renderer");
const renderer = argIndex >= 0 ? process.argv[argIndex + 1] : "deterministic";
const manifest = JSON.parse(
  await readFile(
    new URL("../test/accuracy-manifest.json", import.meta.url),
    "utf8",
  ),
);
const sha = (value) => createHash("sha256").update(value).digest("hex");
function makeCase(entry) {
  const lines = entry.targetPages * 23 - (entry.variant % 3);
  let markdown;
  switch (entry.category) {
    case "plain-wrap":
      markdown = Array.from(
        { length: lines },
        (_, i) =>
          `The record supports conclusion ${i + 1}.${i < lines - 1 ? "  " : ""}`,
      ).join("\n");
      break;
    case "emphasis-heading-list":
      markdown =
        `# Argument ${entry.variant + 1}\n\n` +
        Array.from(
          { length: Math.max(1, lines - 1) },
          (_, i) =>
            `${i % 2 ? "*Authority*" : "**Evidence**"} supports item ${i + 1}.${i < lines - 2 ? "  " : ""}`,
        ).join("\n");
      break;
    case "quote-hard-page-break":
      markdown = Array.from(
        { length: lines },
        (_, i) =>
          `> Quoted authority line ${i + 1}.${i < lines - 1 ? "  " : ""}`,
      ).join("\n");
      if (entry.variant % 2)
        markdown += `\n\n<!-- pagebreak -->\n\nConclusion.`;
      break;
    case "widow-orphan-keep":
      markdown = Array.from(
        { length: Math.max(1, Math.ceil(lines / 4)) },
        (_, i) =>
          `## Point ${i + 1}\n\n` +
          `Supporting legal prose sentence. `.repeat(10),
      ).join("\n\n");
      break;
    case "footnotes": {
      const definitions = [];
      const body = [];
      for (let i = 0; i < lines; i++) {
        if (i % 20 === 0) {
          const id = Math.floor(i / 20) + 1;
          body.push(
            `Body line ${i + 1} with authority.[^${id}]${i < lines - 1 ? "  " : ""}`,
          );
          definitions.push(
            `[^${id}]: Footnote authority for lines beginning at ${i + 1}.`,
          );
        } else
          body.push(
            `Body line ${i + 1} without a new footnote.${i < lines - 1 ? "  " : ""}`,
          );
      }
      markdown = body.join("\n") + "\n\n" + definitions.join("\n");
      break;
    }
    default:
      markdown = Array.from(
        { length: lines },
        (_, i) =>
          `Template geometry line ${i + 1}.${i < lines - 1 ? "  " : ""}`,
      ).join("\n");
  }
  const options =
    entry.category === "template-geometry-style"
      ? {
          layout: {
            page: {
              marginsTwips: {
                top: 1080,
                right: 1260,
                bottom: 1080,
                left: 1260,
              },
            },
            body: {
              fontSizeTwips: entry.variant % 2 ? 260 : 240,
              lineSpacing: { rule: "exact", twips: 480 },
              keepLines: false,
            },
          },
        }
      : {};
  return { markdown, options };
}
if (renderer === "word" && process.env.AGENT_DOCX_TEST_WORD !== "1")
  throw new Error(
    "Set AGENT_DOCX_TEST_WORD=1 to run the live Word release gate",
  );
if (
  renderer === "libreoffice" &&
  process.env.AGENT_DOCX_TEST_LIBREOFFICE !== "1"
)
  throw new Error(
    "Set AGENT_DOCX_TEST_LIBREOFFICE=1 to run the live LibreOffice release gate",
  );
const observations = [];
for (const entry of manifest.cases) {
  const generated = makeCase(entry);
  try {
    const result = await measureMarkdown(generated.markdown, {
      ...generated.options,
      renderer,
    });
    const measured =
      renderer === "word" && result.renderers.word?.status === "ok"
        ? result.renderers.word.value
        : renderer === "libreoffice" &&
            result.renderers.libreoffice?.status === "ok"
          ? result.renderers.libreoffice.value
          : null;
    observations.push({
      id: entry.id,
      category: entry.category,
      targetPages: entry.targetPages,
      markdownSha256: sha(generated.markdown),
      profileSha256: sha(JSON.stringify(result.deterministic.profile)),
      fontSha256: result.deterministic.profile.metricFonts.map((f) => f.sha256),
      deterministicPages: result.deterministic.pageCount,
      rendererPages: measured?.pageCount ?? result.pageCount,
      environment: measured,
    });
    process.stderr.write(
      `${entry.id}: ${result.deterministic.pageCount}/${measured?.pageCount ?? result.pageCount}\n`,
    );
  } catch (error) {
    if (error instanceof AgentDocxError) {
      console.error(`${entry.id}: ${error.code}: ${error.message}`);
      process.exitCode = 4;
      break;
    }
    throw error;
  }
}
if (observations.length === manifest.cases.length) {
  const errors = observations.map((o) =>
    Math.abs(o.deterministicPages - o.rendererPages),
  );
  const exact = errors.filter((n) => n === 0).length / errors.length;
  const mae = errors.reduce((a, b) => a + b, 0) / errors.length;
  const worst = Math.max(...errors);
  const report = {
    schemaVersion: 1,
    renderer,
    caseCount: observations.length,
    exactMatchRate: exact,
    meanAbsolutePageError: mae,
    worstPageError: worst,
    observations,
  };
  console.log(JSON.stringify(report, null, 2));
  if (renderer !== "deterministic" && (exact < 0.95 || mae > 0.1 || worst > 1))
    process.exitCode = 1;
}
