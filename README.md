# agent-docx

**Agent-first legal Markdown-to-DOCX compiler with deterministic page and line constraints, native redlines, embedded revisions, and optional Word validation.**

agent-docx is a local-first drafting suite for litigation documents and court filings. It owns Markdown parsing, legal-document compilation, OOXML package generation, deterministic pagination, revision records, semantic diffs, native tracked changes, and rule-pack validation. It generates DOCX itself—Pandoc and any other Markdown-to-DOCX converter are not required.

> Filing-rule findings are evidence-backed software results, not legal advice or filing certification. Confirm current court, local, judge, document-type, and case-specific requirements before filing.

## Install

Requires Node.js 24 or newer.

```sh
pnpm add agent-docx
agent-docx --version
agent-docx --help
```

## Quick start: a revision-bound filing workflow

Create `metadata.json` with the litigation metadata required by the chosen rule pack, then initialize a project:

```sh
agent-docx project init \
  --project agent-docx.json \
  --document motion \
  --source motion.md \
  --profile cand-civil \
  --filing-kind motion-document \
  --rule-pack cand-civil@2026-05-01 \
  --metadata metadata.json

agent-docx revision checkpoint \
  --project agent-docx.json \
  --document motion \
  --base HEAD \
  --author "Drafter" \
  --message "Initial motion"

agent-docx draft guidance --project agent-docx.json --document motion
agent-docx validate --project agent-docx.json --document motion

agent-docx export \
  --project agent-docx.json \
  --document motion \
  --revision HEAD \
  --mode clean \
  --output motion.docx
```

A project manifest owns document configuration; `.agent-docx/` is a private, content-addressed object store. Revision operations are optimistic and revision-bound: a mutation must identify the expected base revision, and a stale base is rejected instead of silently overwriting another draft.

## Who it is for

- Agent developers building controlled drafting, review, and filing workflows.
- Litigation teams keeping source in Markdown while delivering DOCX.
- Legal operations teams requiring reproducible local artifacts, revision records, and machine-readable diagnostics.
- Lawyers who want optional Microsoft Word or LibreOffice checks without making either application a generation dependency.

## Core workflow

1. **Draft** source-mapped legal Markdown.
2. **Checkpoint** an immutable revision with author and message provenance.
3. **Evaluate** a proposed `SourcePatch` before changing the working copy.
4. **Apply** a verified patch only when its base revision and deterministic checks still match.
5. **Validate** the selected revision against the pinned rule pack and source-mapped structure.
6. **Export** either a clean DOCX or a native-redline DOCX; optionally verify it with a local Office renderer.

`draft guidance` reports page, section, paragraph-tail, last-page, and counted-line budgets. `draft evaluate` returns a canonical patch hash and preflight evidence; `draft apply` requires that hash and can gate on a passing evaluation. This prevents an agent from applying a differently measured patch.

## Commands

```text
agent-docx measure [FILE.md|-] [options]
agent-docx profiles [--json]
agent-docx template inspect FILE.docx [--json]

agent-docx project init|add ...
agent-docx document configure ...
agent-docx revision checkpoint|list|show|restore|diff|resolve ...
agent-docx draft guidance|evaluate|apply ...
agent-docx review add|resolve ...
agent-docx validate ...
agent-docx export ...
agent-docx import ...
agent-docx agent --input-jsonl
agent-docx agent --watch --project FILE --document ID --jsonl
```

All workflow commands accept `--project FILE`; project creation and document addition use `--document`, `--source`, `--profile`, and `--metadata`. `project init` creates the manifest and first document; `project add` adds another document and accepts `--default` to make it the default document. Custom font input requires `--font-family` and `--font-regular` together; bold, italic, and bold-italic files are optional.

```sh
# Inspect a supported template without copying arbitrary package parts.
agent-docx template inspect court-template.docx --json

# Add a review annotation to an immutable revision.
agent-docx review add \
  --project agent-docx.json --document motion --revision HEAD \
  --block blk-argument --start 0 --end 18 \
  --author "Reviewer" --message "Confirm authority"

# See a semantic revision change set.
agent-docx revision diff \
  --project agent-docx.json --document motion BASE_REVISION HEAD

# Export a native tracked-change document. --base identifies the comparison revision.
agent-docx export \
  --project agent-docx.json --document motion --revision HEAD \
  --base BASE_REVISION --mode redline --output motion-redline.docx
```

`revision restore` creates a new revision restoring a target revision; it does not mutate historical records. `revision resolve` applies explicit accept/reject decisions to a change set. `document configure` records configuration changes as a revision. `import --inspect-only` is stateless; normal import requires a target document, output Markdown path, author, and message.

## Agent protocol

`agent-docx agent --input-jsonl` accepts one closed JSON request per stdin line and emits one closed JSON result or error record per accepted line. It is designed for stateful tools rather than shell parsing.

```sh
printf '%s\n' \
  '{"schemaVersion":1,"id":"measure-1","action":"document.measure","project":"agent-docx.json","params":{"documentId":"motion"}}' \
  '{"schemaVersion":1,"id":"validate-1","action":"document.validate","project":"agent-docx.json","params":{"documentId":"motion","revision":"HEAD"}}' \
  | agent-docx agent --input-jsonl
```

Actions include `project.init`, `project.add`, `project.get`, `document.configure`, `document.get`, `document.measure`, `document.validate`, `revision.checkpoint`, `revision.list`, `revision.get`, `revision.restore`, `revision.diff`, `revision.resolve`, `draft.guidance`, `draft.evaluate`, `draft.apply`, `review.add`, `review.resolve`, `docx.export`, `docx.import`, and `docx.inspect`.

The response envelope always contains `schemaVersion`, `kind`, `sequence`, `requestId`, `action`, `project`, `documentId`, and `revision`. Generated binary DOCX bytes are never embedded in CLI JSON; serializable responses provide public artifact paths, SHA-256 values, block manifests, validation, and renderer provenance instead.

For continuous local feedback:

```sh
agent-docx agent --watch \
  --project agent-docx.json --document motion --jsonl
```

The watch stream emits `ready`, debounced `document.measure` results or errors, and `end` on `SIGINT` or `SIGTERM`. Its ready record inventories the source, template, custom fonts, and assets used by the revision. Changing the manifest, source, template, font, an existing asset, or an asset-directory membership refreshes the measurement and watch set.

## Markdown and legal structure

agent-docx supports source-mapped paragraphs, headings, blockquotes, ordered and unordered lists, GFM tables, thematic breaks, hard breaks, emphasis, strong emphasis, strikethrough, inline code, safe absolute links, footnotes, and controlled directives.

Supported directives are deliberately narrow:

- `caption`, `toc`, `toa`, and `pagebreak`.
- `signature` with a known counsel ID; `certificate` with a known certificate ID.
- `sectionbreak` with `kind="next-page"|"continuous"`, plus paired page-number format and start when needed.
- `numbered` paragraphs with a sequence and level 1–4.
- `exhibit` and `image` with project-relative assets; images require positive twip dimensions and alt text.
- `length-exclusion` for the closed set of supported filing-length exclusions.
- Inline legal references and authorities with validated metadata.

The parser rejects arbitrary HTML, YAML front matter, code blocks, math, unsafe or relative ordinary links, remote asset fetching, unknown directives or attributes, and unsupported document structures. It reports source-aware errors rather than silently approximating content.

### Source markers

Legal blocks have stable IDs and source positions. Existing markers are preserved; missing markers can be inserted through the project workflow. This source map drives diagnostics, review annotations, semantic diffs, DOCX bookmarks, redline attribution, and validation evidence.

## Revisions, diffs, and native redlines

A checkpoint stores canonical source, document configuration, resolved dependency object IDs, annotations, a source SHA-256, and a working-tree hash. Revision IDs are content hashes. The store deduplicates immutable objects and writes through locks and atomic swaps.

`revision diff` produces a semantic `ChangeSet`, not a raw line diff. It identifies additions, removals, replacements, moves, configuration changes, dependency changes, and annotation changes using legal-block identities and source ranges. `revision resolve` records explicit decisions against that set.

A redline export converts the resolved semantic diff into native OOXML insertions and deletions, and emits review annotations as native Word comments anchored to their source block. It is a generated comparison artifact—not a substitute for checking the final document in the filing environment. Redline output has revision and comparison provenance, including native tracked-change and comment counts.

Native redline export currently supports body paragraphs, headings, blockquotes, and controlled numbered paragraphs without footnotes. Comments are block-anchored; exact subrange anchors remain source-side review metadata. It rejects lists, tables, images, exhibits, length-exclusion containers, breaks, captions, TOC/TOA fields, and footnotes with `DOCX_REDLINE_UNSUPPORTED` rather than emitting a misleading partial redline. Use clean export for documents containing those constructs.

## Rule packs and validation

Layout profiles control geometry and styles. Rule packs control legal validation. Built-in packs are versioned snapshots:

- `frap-32@2024-12-01`
- `cand-civil@2026-05-01`

Each pack records its official URL, effective date, checked-in source excerpt, SHA-256, exact predicates modeled by the software, and unmodeled provisions. Validation findings include check ID, status, severity, source/evidence, remediation where available, and the rule-pack snapshot. A changing court website cannot silently change existing validation semantics.

The current closed check family covers length alternatives, page size, minimum margins, typeface, line spacing, maximum counted lines, required metadata, required blocks, required footer content, and reference integrity. An `unknown` result means the deterministic compiler cannot establish the relevant native behavior; it is not a passing filing result.

## Layout, pagination, and profiles

The default `us-district-conventional` profile is a product baseline: U.S. Letter, one-inch margins, 12-point Times-compatible serif metrics, and double spacing. It is not a filing certification. `frap-32` and `cand-civil` add source metadata and filing constraints.

```sh
agent-docx profiles
agent-docx profiles --json
agent-docx measure filing.md --profile cand-civil --paragraphs --sections --trim
```

Portable estimates use pinned Liberation Serif 2.1.5 bytes as metrics while reporting the requested `Times New Roman` family and explicit substitution. Provide legally obtained custom font files through project configuration when a different deterministic metric source is required.

Deterministic pagination runs entirely in process:

1. Markdown is normalized into source-mapped legal blocks and footnotes.
2. Profile, supported template input, and configuration determine geometry, indentation, styles, line spacing, and page rules.
3. `fontkit` shapes runs with pinned font bytes; Unicode line breaking wraps text to usable width.
4. The paginator places lines, table rows, paragraph spacing, keeps, widow/orphan controls, explicit breaks, and counted-line caps.
5. Footnotes reserve bottom-page space when first referenced and report a relaxed constraint if an intrinsic split is unavoidable.
6. Results report physical pages, fractional equivalent-page use, visual and counted lines, paragraph-tail diagnostics, section attribution, and last-page metrics.

The same Markdown, configuration, and metric-font bytes produce the same deterministic result. The model reports unsupported content instead of guessing.

## DOCX templates and import

`template inspect` uses a bounded ZIP/XML reader to inspect supported style inheritance, numbering, theme/font information, sections, header/footer relationships, fields, and caption components. Unsafe or unsupported package features—including macros, external relationships, embedded objects, scripts, encrypted parts, and arbitrary package copying—are not executed or copied into output.

Supported inspected layout/style data can be consumed by project configuration. A template is input for supported style and geometry semantics, not a host document to merge into an output package.

`import` reads supported DOCX material into legal Markdown and returns fidelity classifications: `preserved`, `normalized`, `externalized`, or `unsupported`. Import is explicit about information it cannot preserve; inspect-only import does not write a project revision.

## Optional Microsoft Word and LibreOffice validation

The default renderer is `deterministic`: it neither discovers nor starts Office applications. `word`, `libreoffice`, and `compare` are explicit local opt-ins after agent-docx generates its DOCX.

```sh
agent-docx measure filing.md --renderer word --json
agent-docx measure filing.md --renderer libreoffice --json
agent-docx export \
  --project agent-docx.json --document motion --revision HEAD \
  --mode clean --renderer compare --output motion.docx
```

- **Word** runs a hardened, noninteractive Windows PowerShell/COM bridge with macros disabled, a local mutex, absolute executable paths, temporary DOCX input, and generated bookmark diagnostics.
- **LibreOffice** runs headlessly with an isolated user profile, fixed locale and timezone, and validates a generated PDF page tree. It is never labeled as Word.
- **Compare** retains the deterministic result as the stable top-level baseline and records each successful Office result or structured error.

Office versions, installed fonts, active printer, and layout engines can change pagination. Those applications are useful validation engines, not hidden dependencies or portable goldens. Native behaviors agent-docx cannot verify without Office remain explicitly unknown or warned rather than blocking deterministic output.
The corpus-level Word parity check is opt-in: set `AGENT_DOCX_TEST_WORD=1` when running the native comparison. The normal `pnpm test` suite does not require Microsoft Word or LibreOffice.

## JavaScript API

```ts
import {
  createProject,
  openProject,
  compileMarkdown,
  estimateMarkdown,
  measureMarkdown,
  inspectDocxTemplate,
} from "agent-docx";

const estimate = await estimateMarkdown(markdown, {
  profile: "cand-civil",
  pageLimit: 25,
  paragraphDiagnostics: true,
});

const compiled = await compileMarkdown(markdown, {
  documentId: "motion",
  profile: "cand-civil",
  metadata,
});

const project = await createProject("agent-docx.json", {
  documentId: "motion",
  source: "motion.md",
  profile: "cand-civil",
  metadata,
});
const state = await project.getState();
```

`estimateMarkdown` is portable deterministic pagination. `measureMarkdown` always computes that estimate and can optionally invoke Office. `compileMarkdown` returns generated DOCX bytes, a block manifest, validation, and deterministic measurement. `createProject` and `openProject` provide revision-bound project operations. See exported TypeScript declarations for complete option and result types.

## Published machine contracts

All published schemas are JSON Schema Draft 2020-12 and are exported from the package:

- `agent-docx/project.schema.json`
- `agent-docx/rule-pack.schema.json`
- `agent-docx/revision.schema.json`
- `agent-docx/change-set.schema.json`
- `agent-docx/source-patch.schema.json`
- `agent-docx/validation-result.schema.json`
- `agent-docx/artifact-result.schema.json`
- `agent-docx/compiled-docx.schema.json`
- `agent-docx/docx-import-result.schema.json`
- `agent-docx/agent-request.schema.json`
- `agent-docx/agent-response.schema.json`
- `agent-docx/agent-stream.schema.json`
- `agent-docx/measurement-request.schema.json`
- `agent-docx/measurement-result.schema.json`
- `agent-docx/measurement-stream.schema.json`
- `agent-docx/docx-template-inspection.schema.json`
- `agent-docx/profile-catalog.schema.json`
- `agent-docx/config.schema.json`

Schemas are closed where a protocol or stored record needs a stable contract. The generic JSON value definitions intentionally permit valid JSON payload content only where the format requires extensibility.

## Security and local data handling

- Project source, template, font, and asset paths must be contained within the project and resolve to regular non-symlink files.
- Dependency bytes are stored by SHA-256 in `.agent-docx`; immutable revisions refer to those objects.
- DOCX ZIP/XML reading enforces entry, compressed/uncompressed-size, expansion-ratio, path, UTF-8, and entity/DTD boundaries.
- Generation does not call a shell or fetch remote content.
- Word and LibreOffice run only when requested. Their paths are explicit or hardened discovery results, and their outputs are provenance rather than source of truth.
- CLI machine output never contains binary DOCX data or internal absolute paths.

## Development

```sh
pnpm test
pnpm verify:pack
pnpm social:preview
```

`pnpm social:preview` deterministically regenerates `docs/assets/agent-docx-social-preview.png`. Launch copy, release template, social posts, repository metadata, social-preview alt text, and upload instructions live in [`docs/marketing-kit.md`](docs/marketing-kit.md).

## License

MIT. See [LICENSE](LICENSE). Third-party notices are in [THIRD_PARTY_NOTICES.txt](THIRD_PARTY_NOTICES.txt).
