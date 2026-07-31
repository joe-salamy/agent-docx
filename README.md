# md-page-count

Deterministic DOCX-equivalent pagination for legal-prose Markdown, with opt-in measurements from locally installed Microsoft Word and LibreOffice Writer.

## Install and run

Requires Node.js 24 or newer.

```sh
pnpm add md-page-count
md-page-count --list-profiles
md-page-count filing.md
md-page-count filing.md --profile frap-32 --paragraphs --trim
md-page-count filing.md --renderer compare --json
```

Standard input is accepted by omitting the file or passing `-`. Human-readable output is the default. Use `--json` for one machine-readable result, `--list-profiles` to discover bundled profiles, `--output FILE.docx` to export the exact generated package for a single input, `--batch INPUT...` for ordered JSONL, `--watch FILE --jsonl` for change streams, and `--inspect-template FILE.docx --json` for bounded, read-only template inspection. Positional batch inputs may be explicit files, directories, or glob patterns; directory discovery is recursive with `*.md` by default and supports repeatable `--include`/`--exclude` patterns plus `--no-recursive`.

## Agent JSON and JSONL protocol

Single-input `--json` writes one unwrapped version-1 `MeasurementResult`; `--inspect-template --json` writes one unwrapped `DocxTemplateInspection`; `--list-profiles --json` writes one unwrapped profile catalog. Batch and watch streams use one version-1 JSON object per line. Successful machine modes keep stderr empty.

`--batch --input-jsonl` accepts file and inline requests:

```jsonl
{"id":"motion","path":"drafts/motion.md"}
{"id":42,"name":"reply","markdown":"# Reply\n\nArgument."}
```

Each request is closed: unknown keys, an empty path, invalid IDs, and requests containing both or neither `path` and `markdown` are rejected. Shared CLI/config layout, renderer, diagnostics, trim, and limit options apply to every batch item. Result and per-item error records preserve a valid request `id` as `requestId` and any derivable source. Public paths are `/`-separated and relative to the invocation directory, including `../` paths for external dependencies; internal absolute paths are never emitted.
`--batch --input-jsonl` decodes and processes each complete line as it arrives, so pipeline memory is bounded by the largest request rather than the entire stream.

JSONL records are discriminated by `kind`: `result`, `error`, `ready`, or `end`. Every invocation starts `sequence` at 1 and increments it in emission order. Batch records have `mode:"batch"` and `trigger:null`. Watch JSONL emits `ready`, the initial `result` or `error`, source/dependency update records, then `end` on `SIGINT` or `SIGTERM`; watch `requestId` is always `null`, and update records include a source-relative trigger. Fatal setup or argument failures write exactly one `kind:"fatal"` record to stderr.

`--trim` exposes deterministic, heuristic width measurements, not generated edits: candidates are eligible soft-wrapped tails ranked by ascending `oneLineReduction.estimatedRemovalTwips`, with normalized UTF-16 text ranges and exact-or-node source provenance. Re-run pagination after editing.

Exit statuses are 0 for success, 1 for processing/internal failure, 2 for invalid arguments, 3 for `--fail-over-limit`, 4 for Office renderer failure, 130 for `SIGINT`, and 143 for `SIGTERM`.

Runtime Draft 2020-12 schemas are exported as `md-page-count/cli-request.schema.json`, `md-page-count/measurement-result.schema.json`, `md-page-count/docx-template-inspection.schema.json`, `md-page-count/cli-jsonl.schema.json`, `md-page-count/cli-error.schema.json`, and `md-page-count/profile-catalog.schema.json`. Configuration remains `md-page-count/config.schema.json`.

## JavaScript API

```ts
import {
  estimateMarkdown,
  inspectDocxTemplate,
  measureMarkdown,
} from "md-page-count";

const estimate = await estimateMarkdown(markdown, {
  profile: "cand-civil",
  pageLimit: 25,
  paragraphDiagnostics: true,
});

const measured = await measureMarkdown(markdown, {
  profile: "cand-civil",
  renderer: "compare",
  includeGeneratedDocx: true,
});
```

`estimateMarkdown` runs only the portable deterministic paginator. `measureMarkdown` always computes that estimate and optionally invokes an Office renderer. Both return structured provenance, warnings, physical page count, fractional equivalent-page usage, total and per-page visual-line counts, and last-page metrics; page-limit options also add budget information. `sectionDiagnostics: true` adds inclusive, source-ordered heading-section page and line attribution under the deterministic result. `includeGeneratedDocx: true` adds the generated `Uint8Array`; Office renderers consume that same buffer. The CLI never serializes these bytes into JSON. Exported TypeScript types describe the complete result.

## Profiles, fonts, and templates

The default `us-district-conventional` profile is a product baseline: U.S. Letter, one-inch margins, 12-point Times-compatible serif metrics, and double spacing. It does not certify filing compliance. Verify current court, local, judge, document-type, and case-specific rules. The `frap-32` and `cand-civil` profiles carry source citations and effective dates.

Discover profile IDs, source metadata, requested font families, and built-in filing limits without loading the JavaScript API:

```sh
md-page-count --list-profiles
md-page-count --list-profiles --json
```

Portable estimates use pinned Liberation Serif 2.1.5 bytes as metrics while reporting `Times New Roman` as the requested family and the substitution explicitly. Supply legally obtained font bytes with `fontSet` or `--font-regular` and the related font flags when a different deterministic metric source is required. Missing bold or italic faces are reused and reported.

Layout can be adjusted through API overrides or CLI flags for page size, dimensions, margins, body font size, and line spacing. `--template FILE.docx` inspects a template and uses its supported final-section geometry and styles as layout input; it does not execute macros or embed the source document. `inspectDocxTemplate` exposes that inspection separately.

API/config layout overrides also expose `thematicBreak` spacing/thickness/keep behavior and `table` header/body styles, cell padding, border width, and repeated-header behavior. Built-in profiles derive these from the body style; DOCX template inspection does not infer unmodeled table or rule properties.

## How page calculation works

Without an Office opt-in, pagination is calculated in-process and does not create or open a document in desktop software:

1. Markdown is normalized into source-mapped text, table, thematic-break, and page-break blocks plus an ordered map of paragraph-only footnote definitions.
2. The selected profile, template inspection, and overrides determine page geometry, indents, paragraph spacing, font sizes, line spacing, and pagination rules. Usable width and height are the page dimensions minus margins and gutter.
3. `fontkit` shapes each text run with the selected deterministic font bytes. The paginator uses the resulting glyph advance widths and Unicode line-break opportunities to wrap text into the available width; bold and italic runs use their corresponding font faces.
4. Each wrapped line receives a height from the font's ascent, descent, and line-gap metrics, adjusted by the configured line-spacing rule. Lines and paragraph spacing are placed vertically until the usable height or a profile-specific counted-line cap is reached. Feasible keep-with-next chains move together; keep-lines, widow/orphan, and explicit page-break rules can also move content to the next page.
5. A footnote definition reserves bottom space on the page containing its first reference. Repeated references do not reserve it again. An intrinsically oversized definition continues at the bottom of following pages and reports `FOOTNOTE_SPLIT_CONSTRAINT_RELAXED`; the deterministic model does not claim pixel-identical native separator metrics.
6. `pageCount` is the number of physical pages produced. `equivalentPages` adds the full pages before the last page to the fraction of usable height consumed on the last page: `full pages + last-page used height / usable height`.

This is a deterministic layout model, not a hidden Word or Writer process. The same Markdown and resolved styles produce the same result when the metric font bytes are unchanged. It models the supported DOCX layout features directly and reports unsupported content rather than guessing.

With `word`, `libreoffice`, or `compare`, this deterministic calculation still runs first. The library then generates a DOCX from the normalized blocks and resolved styles. Word measures that DOCX with Word's own pagination engine; Writer converts it with Writer's engine and the exported PDF page count is measured. Those applications can differ from the deterministic model and from each other because their layout engines, installed fonts, versions, and environment differ.

## Microsoft Word and LibreOffice measurements

Office rendering is opt-in. The default renderer is `deterministic`, which neither discovers nor starts Word or LibreOffice. For every opt-in mode, the library first resolves the same profile and Markdown flow used by the deterministic paginator, generates a temporary DOCX, and then passes that DOCX to the selected local application. These measurements are renderer- and machine-specific, not replacements for filing-rule review.

Select a mode with the CLI or API:

```sh
md-page-count filing.md --renderer word --json
md-page-count filing.md --renderer libreoffice --json
md-page-count filing.md --renderer compare --json
```

```ts
await measureMarkdown(markdown, { renderer: "word" });
await measureMarkdown(markdown, {
  renderer: "libreoffice",
  libreoffice: { executablePath: "/usr/bin/soffice" },
  officeTimeoutMs: 90_000,
});
```

The modes differ as follows:

- `word`: returns Word's page count as the top-level `pageCount` and sets `pageCountSource` to `word`. Word must render successfully; it never falls back to another renderer.
- `libreoffice`: returns Writer's page count as the top-level value and source. Writer must render successfully; it never falls back to Word or the estimate.
- `compare`: attempts Word and Writer, records each result or structured error under `renderers`, and leaves the top-level page count sourced from the deterministic estimate so comparisons have a stable baseline. It fails only if neither Office renderer succeeds.

### Microsoft Word

The Word adapter requires desktop Microsoft Word on Windows. It also works from WSL by invoking Windows PowerShell through the mounted Windows installation. A custom `word.powerShellPath` can be supplied through the API or config and must be absolute.

The adapter:

1. sends the generated DOCX to a bundled noninteractive PowerShell script;
2. serializes concurrent runs with a local mutex and starts a hidden Word COM instance with macros disabled;
3. verifies that each requested font family appears in Word's installed-font list;
4. opens the DOCX read-only, forces repagination, and reads Word's page count plus total, per-page, and last-page body-line statistics;
5. when paragraph diagnostics or trim analysis is requested, maps generated body-paragraph bookmarks back to source positions and reports each paragraph's Word line count, page span, and final rendered line text; and
6. closes Word and deletes the temporary document.

The result includes Word's line statistics, Word version/build, active printer, requested fonts, generated-DOCX SHA-256, duration, and cleanup state. Requested paragraph measurements appear under `renderers.word.value.paragraphDiagnostics`. Paragraph extraction is best-effort: a bookmark or extraction failure is reported there as a nested renderer error while a valid Word page-count summary remains usable. Word pagination can vary with the installed Word build, fonts, and active printer, which is why that provenance is retained. The default timeout is 120 seconds. Generated DOCX input to this adapter is limited to 25 MiB.

### LibreOffice Writer

The Writer adapter requires a local LibreOffice executable. It discovers `soffice`/`libreoffice` on `PATH` on Linux, the standard application path on macOS, and standard Program Files paths on Windows. Override discovery with `--libreoffice-path` or `libreoffice.executablePath`; explicit API paths must be absolute.

Writer runs headlessly with a fresh isolated user profile, fixed `C` locale, and UTC timezone. It converts the generated DOCX to PDF with Writer's PDF export filter; the library validates that PDF and counts its page tree. Temporary input, output, and profile directories are removed afterward. The result includes the LibreOffice version, executable/platform/architecture, requested fonts, DOCX and PDF SHA-256 values, duration, and font-environment calibration flag. The default timeout is 60 seconds.

LibreOffice does not guarantee the requested font is installed and may substitute another font. The optional `libreoffice.installedFonts` config is provenance supplied by the caller; it marks the environment as calibrated but does not install fonts or independently verify those files. Without it, the result carries `LIBREOFFICE_FONT_ENVIRONMENT_UNVERIFIED`. A Writer measurement is deliberately never labeled as a Word measurement.

`officeTimeoutMs` / `--office-timeout` overrides the render timeout for either adapter and accepts 1,000 through 600,000 milliseconds. Office-mode failures use structured codes such as `WORD_NOT_FOUND`, `WORD_TIMEOUT`, `LIBREOFFICE_NOT_FOUND`, and `LIBREOFFICE_RENDER_FAILED`; the CLI exits with status 4 for renderer availability or execution failures.

## Accuracy testing

`pnpm test` includes unit and boundary coverage plus deterministic golden tests over six committed, real-world Markdown lecture-note documents. All corpus cases invoke the CLI with `--paragraphs`. Three documents were used while correcting pagination; three distinct documents were selected and hash-locked without reading their contents before a one-shot Word comparison. When desktop Word is available, `pnpm test` also invokes `--renderer word` and requires exact physical-page and last-page-line agreement across all six documents; hosts without Word report that live test as skipped.

`pnpm accuracy` runs all six documents together with the synthetic boundary matrix. Pass `--renderer word` or `--renderer libreoffice` with the corresponding `MD_PAGE_COUNT_TEST_WORD=1` or `MD_PAGE_COUNT_TEST_LIBREOFFICE=1` opt-in to compare the same inputs with a native renderer. The release gate evaluates exact-match rate, mean absolute page error, and worst page error across the corpus; native output is not a portable per-document golden because Office versions, fonts, and printer state can change pagination.

## Supported Markdown

Supported block content includes paragraphs, headings, blockquotes, ordered and unordered lists, GFM tables, thematic breaks, paragraph-only multi-paragraph GFM footnotes, and the explicit `<!-- pagebreak -->` marker. Inline text may use emphasis, strong emphasis, links, strikethrough, hard breaks, same-style literal inline code, and footnote references. Table cells support those text styles and GFM alignment, except footnote references. Link destinations do not affect layout.

Code blocks, images, arbitrary HTML, YAML front matter, math, non-paragraph footnote children, and block/media content in table cells are rejected with source-aware `UNSUPPORTED_MARKDOWN` instead of being silently approximated.

## Diagnostics and configuration

Use `--paragraphs` for per-paragraph line and last-line-fill diagnostics. `--sections` reports inclusive heading-section page and line usage; with a page limit it also identifies each section touching pages beyond that limit. `--trim` reports advisory candidates whose final line is short; `--trim-limit` and `--trim-threshold` tune that report. `--page-limit` adds remaining/over-limit budget data, and `--fail-over-limit` makes an over-limit CLI result fail.

Deterministic paragraph diagnostics include final-line text, normalized UTF-16 range, exact-or-node source ranges, used/available/unused widths, overflow state, and the previous line. `oneLineReduction` is present only for eligible multi-line soft wraps and reports the estimated width that must be removed or rephrased; it is heuristic, never an edit. Human `--trim` output ranks these candidates by that width. `--paragraphs --trim` additionally shows the generic fill bars.

Pass `--config path/to/config.json` to use a JSON configuration file. Paths inside it are resolved relative to that file; explicit CLI options override config values. Top-level `sectionDiagnostics` enables section output, while `batch.recursive`, `batch.include`, and `batch.exclude` configure positional batch discovery. Configuration is validated against the exported `md-page-count/config.schema.json` schema. No parent-directory, home-directory, or environment configuration is discovered.

Widow/orphan control is enabled by default. Set `layout.pagination.widowOrphanControl` to `false` to disable it in both deterministic pagination and generated DOCX paragraphs; `widowLines` and `orphanLines` continue to configure the enabled minimums and must each be integers of at least 1.
