import { readFile } from "node:fs/promises";
import { createHash } from "node:crypto";
import { measureMarkdown, AgentDocxError } from "../dist/index.js";
const argIndex = process.argv.indexOf("--renderer");
const renderer = argIndex >= 0 ? process.argv[argIndex + 1] : "deterministic";
const syntheticManifest = JSON.parse(
  await readFile(
    new URL("../test/accuracy-manifest.json", import.meta.url),
    "utf8",
  ),
);
const corpusManifest = JSON.parse(
  await readFile(
    new URL("../test/corpus-manifest.json", import.meta.url),
    "utf8",
  ),
);
const blindCorpusManifest = JSON.parse(
  await readFile(
    new URL("../test/blind-corpus-manifest.json", import.meta.url),
    "utf8",
  ),
);
const realCases = (manifest, directory) =>
  manifest.documents.map((document) => ({
    id: document.id,
    category: "real-document",
    targetPages: document.expected.pageCount,
    fixture: `${directory}/${document.file}`,
    markdownSha256: document.markdownSha256,
  }));
const cases = [
  ...syntheticManifest.cases,
  ...realCases(corpusManifest, "corpus"),
  ...realCases(blindCorpusManifest, "blind-corpus"),
];
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
async function loadCase(entry) {
  if (!entry.fixture) return makeCase(entry);
  const markdown = await readFile(
    new URL(`../test/fixtures/${entry.fixture}`, import.meta.url),
    "utf8",
  );
  if (sha(markdown) !== entry.markdownSha256)
    throw new Error(
      `${entry.id}: corpus fixture SHA-256 does not match manifest`,
    );
  return { markdown, options: {} };
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
const rendererOptions =
  renderer === "word" && process.env.AGENT_DOCX_ACCURACY_WORD_PATH
    ? {
        word: {
          powerShellPath: process.env.AGENT_DOCX_ACCURACY_WORD_PATH,
        },
      }
    : renderer === "libreoffice" &&
        process.env.AGENT_DOCX_ACCURACY_LIBREOFFICE_PATH
      ? {
          libreoffice: {
            executablePath: process.env.AGENT_DOCX_ACCURACY_LIBREOFFICE_PATH,
          },
        }
      : {};
const pageMetrics = (errors) => ({
  exactMatchRate: errors.filter((error) => error === 0).length / errors.length,
  meanAbsolutePageError:
    errors.reduce((total, error) => total + error, 0) / errors.length,
  worstPageError: Math.max(...errors),
});
if (!["deterministic", "word", "libreoffice"].includes(renderer))
  throw new Error(
    `Accuracy accepts --renderer deterministic, word, or libreoffice (received ${renderer})`,
  );
const observations = [];
for (const entry of cases) {
  const generated = await loadCase(entry);
  try {
    const result = await measureMarkdown(generated.markdown, {
      ...generated.options,
      ...rendererOptions,
      renderer,
    });
    const rendererRecord =
      renderer === "word"
        ? result.renderers.word
        : renderer === "libreoffice"
          ? result.renderers.libreoffice
          : undefined;
    let measured;
    if (renderer !== "deterministic") {
      if (!rendererRecord)
        throw new Error(
          `requested renderer ${renderer} did not return a status`,
        );
      if (rendererRecord.status !== "ok")
        throw new AgentDocxError(
          rendererRecord.error.code,
          `${renderer} renderer ${rendererRecord.status}: ${rendererRecord.error.message}`,
        );
      measured = rendererRecord.value;
    }
    observations.push({
      id: entry.id,
      category: entry.category,
      targetPages: entry.targetPages,
      markdownSha256: sha(generated.markdown),
      profileSha256: sha(JSON.stringify(result.deterministic.profile)),
      fontSha256: result.deterministic.profile.metricFonts.map((f) => f.sha256),
      deterministicPages: result.deterministic.pageCount,
      ...(measured
        ? { rendererPages: measured.pageCount, environment: measured }
        : {}),
    });
    process.stderr.write(
      `${entry.id}: deterministic=${result.deterministic.pageCount}, target=${entry.targetPages}` +
        (measured ? `, ${renderer}=${measured.pageCount}` : "") +
        "\n",
    );
  } catch (error) {
    const code = error instanceof AgentDocxError ? error.code : "UNKNOWN";
    const message = error instanceof Error ? error.message : String(error);
    console.error(
      `${entry.id}: requested ${renderer} renderer failed (${code}): ${message}`,
    );
    process.exitCode = 4;
    break;
  }
}
if (observations.length === cases.length) {
  const deterministic = pageMetrics(
    observations.map((observation) =>
      Math.abs(observation.deterministicPages - observation.targetPages),
    ),
  );
  const rendererMetrics =
    renderer === "deterministic"
      ? deterministic
      : pageMetrics(
          observations.map((observation) =>
            Math.abs(observation.rendererPages - observation.targetPages),
          ),
        );
  const report = {
    schemaVersion: 1,
    renderer,
    caseCount: observations.length,
    exactMatchRate: rendererMetrics.exactMatchRate,
    meanAbsolutePageError: rendererMetrics.meanAbsolutePageError,
    worstPageError: rendererMetrics.worstPageError,
    ...(renderer === "deterministic" ? {} : { deterministic }),
    observations,
  };
  console.log(JSON.stringify(report, null, 2));
  if (
    renderer !== "deterministic" &&
    (rendererMetrics.exactMatchRate < 0.95 ||
      rendererMetrics.meanAbsolutePageError > 0.1 ||
      rendererMetrics.worstPageError > 1)
  )
    process.exitCode = 1;
}
