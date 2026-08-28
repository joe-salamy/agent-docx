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

The package ships versioned agent skills under `skills/` (`skills/agent-docx` for the full drafting workflow, `skills/brief-to-agent-docx` as a thin alias). Install them into the consuming repo as an opt-in step — no postinstall side effect:

```sh
agent-docx skills list
agent-docx skills install              # -> ./.omp/skills (oh-my-pi)
agent-docx skills install --dest .claude/skills  # Claude Code
agent-docx skills install --global     # -> ~/.omp/skills
agent-docx skills install --dry-run --json  # preview
```

Re-run after `pnpm update agent-docx` to pick up skill updates; use `--force` to overwrite an existing install. Skills are part of the published `files` and available at `node_modules/agent-docx/skills/` without copying.

### Input and package limits

The CLI rejects oversized inputs before parsing: Markdown and stdin are capped at 64 MiB, and each JSONL request line is capped at 8 MiB. DOCX packages are capped at 25 MiB compressed input, 512 ZIP entries, 64 MiB decompressed total, and 4 MiB per XML part (with a 12 MiB XML total); attachment bundles allow at most 512 entries, 25 MiB per file, and 50 MiB decompressed total. These limits protect both the CLI and long-running agent transports; a rejected input is not partially processed.

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
agent-docx skills list [--json]
agent-docx skills install [--dest DIR] [--global] [--force] [--dry-run] [--json]

agent-docx project init|add ...
agent-docx document configure ...
agent-docx revision checkpoint|list|show|restore|diff|resolve ...
agent-docx draft guidance|evaluate|apply ...
agent-docx review add|resolve ...
agent-docx validate ...
agent-docx export ...
agent-docx import ...
agent-docx import-redline ...
agent-docx filing-set add|remove|get|validate ...
agent-docx agent --input-jsonl
agent-docx agent --watch --project FILE --document ID --jsonl
agent-docx mcp
```

All workflow commands accept `--project FILE`; project creation and document addition use `--document`, `--source`, `--profile`, and `--metadata`. `project init` creates the manifest and first document; `project add` adds another document and accepts `--default` to make it the default document. Font flags are an all-or-nothing family set: `--font-family` and `--font-regular` are required together, while `--font-bold`, `--font-italic`, and `--font-bold-italic` are optional face overrides. Supplying any face flag without `--font-regular` is rejected.

### Filing sets

Documents can be grouped into ordered filing sets with an optional shared deterministic page cap:

```sh
agent-docx filing-set add \
  --project agent-docx.json --id motion-package \
  --label "Motion package" --documents motion,brief,proposed-order \
  --page-cap 30

agent-docx filing-set get --project agent-docx.json --id motion-package
agent-docx filing-set validate --project agent-docx.json --id motion-package
agent-docx filing-set remove --project agent-docx.json --id motion-package
```

Filing sets live in the project manifest (`filingSets`), reference existing documents in order, and are validated at add time (unknown or duplicate documents are rejected). `filing-set validate` (agent actions `filingSet.add`/`remove`/`get`/`validate`) reports each member document's validation result and deterministic page count, and when a `pageCap` is set, the summed page budget with `pass`/`fail`/`unknown` status—an unmeasured member makes the budget `unknown`, never a silent pass. Set membership is manifest state, not a revision; document revisions stay independently revision-bound.

The published project schema enforces nonempty relative IDs and unique array items. JSON Schema cannot compare object properties across array elements, so duplicate document IDs with different configurations and filing-set `documentIds` references remain explicit runtime manifest invariants.

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

`agent-docx agent --input-jsonl` accepts one closed JSON request per stdin line and emits one closed JSON result or error record per accepted line. It is designed for stateful tools rather than shell parsing. `--input-jsonl` is an agent transport mode and is rejected when combined with `--batch`; use either the JSONL request stream or batch file discovery, never both.

```sh
printf '%s\n' \
  '{"schemaVersion":1,"id":"measure-1","action":"document.measure","project":"agent-docx.json","params":{"documentId":"motion"}}' \
  '{"schemaVersion":1,"id":"validate-1","action":"document.validate","project":"agent-docx.json","params":{"documentId":"motion","revision":"HEAD"}}' \
  | agent-docx agent --input-jsonl
```

Actions include `project.init`, `project.add`, `project.get`, `document.configure`, `document.get`, `document.measure`, `document.validate`, `revision.checkpoint`, `revision.list`, `revision.get`, `revision.restore`, `revision.diff`, `revision.resolve`, `draft.guidance`, `draft.evaluate`, `draft.apply`, `review.add`, `review.resolve`, `docx.export`, `docx.import`, `docx.inspect`, `docx.importRedline`, `filingSet.add`, `filingSet.remove`, `filingSet.get`, and `filingSet.validate`.

The response envelope always contains `schemaVersion`, `kind`, `sequence`, `requestId`, `action`, `project`, `documentId`, and `revision`. Generated binary DOCX bytes are never embedded in CLI JSON; serializable responses provide public artifact paths, SHA-256 values, block manifests, validation, and renderer provenance instead.

For continuous local feedback:

```sh
agent-docx agent --watch \
  --project agent-docx.json --document motion --jsonl
```

The watch stream emits `ready`, debounced `document.measure` results or errors, and `end` on `SIGINT` or `SIGTERM`. Its ready record inventories the source, template, custom fonts, and assets used by the revision. Changing the manifest, source, template, font, an existing asset, or an asset-directory membership refreshes the measurement and watch set.

## Model Context Protocol server

```sh
agent-docx mcp
```

`agent-docx mcp` serves the same version-1 protocol as a Model Context Protocol server over stdio (newline-delimited JSON-RPC, no framing). Every protocol action becomes one MCP tool named after the action (for example `document.validate`, `draft.evaluate`, `docx.export`, `filingSet.get`); each tool takes the action's `params` object plus an optional `project` path relative to the server's working directory. Tool results carry the serialized protocol value as both text and `structuredContent`; dispatch failures return `isError: true` results with a `{code, message}` structured payload. The server implements `initialize`, `ping`, `tools/list`, and `tools/call`, so MCP-capable agents (Claude Code, Cursor, and similar) can drive the full project, draft, review, validation, and export workflow without shell parsing.
MCP project paths are confined to the server working directory. Requests may name a project with a normalized relative path, but absolute paths and traversal outside that directory are rejected; responses expose cwd-relative public paths. Run the server from a directory that contains only the projects and assets you intend to make available.

## Agent skills

The package ships versioned skills as the canonical agent interface — no copy-paste of workflow steps:

- `skills/agent-docx` — full drafting workflow (measure, project/revision/draft/validate/export/filing-set/redline/import, deterministic pagination, rule packs, agent JSONL, and MCP). This is the skill to install.
- `skills/brief-to-agent-docx` — thin alias for the common brief/motion → DOCX path.

Skills are part of the published `files` and live at `node_modules/agent-docx/skills/` after `pnpm add agent-docx`. The CLI copies them into the consumer repo on demand (see Install):

```sh
agent-docx skills list --json
agent-docx skills install --dry-run --json
```

Full instruction text is in `skills/agent-docx/SKILL.md` (versioned with the package); `README.md` keeps only the distro commands. Re-run `skills install` after updating the package; use `--force` to overwrite.

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

### Redline round trip

A redline returned from a reviewer can be imported back into the revision store as an explicit decision set:

```sh
agent-docx import-redline \
  --project agent-docx.json --document motion \
  --input motion-redline-reviewed.docx \
  --author "Drafter" --message "Reviewer accepted all"

agent-docx revision resolve \
  --project agent-docx.json --document motion \
  --change-set change-set.json --decisions decisions.json \
  --author "Drafter" --message "Apply reviewer decisions"
```

`import-redline` (`docx.importRedline`) reads a redline exported by agent-docx whose tracked changes the reviewer resolved in Word (accept or reject; a fully unresolved redline yields an empty decision set for the agent to decide later). It validates the DOCX's semantic manifest against the project, requires the redline revision to be the current HEAD with a clean working copy, and returns a canonical `ChangeSet` plus a `decisions` map covering every change. Reviewer-added Word comments anchored to document blocks become open review annotations. Import either classifies every change unambiguously or fails with `DOCX_IMPORT_UNSUPPORTED`; it never guesses. The returned `ChangeSet` is byte-identical to what `revision.resolve` re-derives, so the decisions commit exactly the reviewer's accepted state as a new revision.

## Rule packs and validation

Layout profiles control geometry and styles. Rule packs control legal validation. Built-in packs are versioned snapshots:

- `frap-32@2024-12-01`
- `cand-civil@2026-05-01`

Each pack records its official URL, effective date, checked-in source excerpt, SHA-256, exact predicates modeled by the software, and unmodeled provisions. Validation findings include check ID, status, severity, source/evidence, remediation where available, and the rule-pack snapshot. A changing court website cannot silently change existing validation semantics.

The current closed check family covers length alternatives, page size, minimum margins, typeface, line spacing, maximum counted lines, required metadata, required blocks, required footer content, and reference integrity. An `unknown` result means the deterministic compiler cannot establish the relevant native behavior; it is not a passing filing result.

### Authorable rule packs

Beyond the built-in snapshots, projects can attach user-defined rule packs: schema-validated JSON files that express the same closed check family as data-driven parameters. A pack file declares an id, source citation, effective date, source excerpt and SHA-256, unmodeled provisions, and a `checks` array whose entries pair a check id with a kind and `params` (for example `page-size` with `{widthTwips, heightTwips}`, `counted-lines-maximum` with `{perPageMaximum}`, or `required-footer` with `{requiredTokens}`). Packs attach to a document through project configuration:

```sh
agent-docx document configure \
  --project agent-docx.json --document motion \
  --base HEAD --changes changes.json \
  --author "Drafter" --message "Attach firm pack"

# changes.json: { "rulePacks": ["packs/firm-style.json"] }
```

Pack files are project-relative, content-hashed into the revision's dependency objects, and schema-validated; a changed pack file changes the working-tree hash, and validation reports any pack whose content no longer matches its snapshot. Custom pack findings appear alongside built-in findings in `ValidationResult`, with each pack recorded in `scope.sourceSnapshots`. Validation semantics for user packs are the user's responsibility; the built-in packs remain evidence-backed snapshots.

## Layout, pagination, and profiles

The default `us-district-conventional` profile is a product baseline: U.S. Letter, one-inch margins, 12-point Times-compatible serif metrics, and double spacing. It is not a filing certification. `frap-32` and `cand-civil` add source metadata and filing constraints.

```sh
agent-docx profiles
agent-docx profiles --json
agent-docx measure filing.md --profile cand-civil --paragraphs --sections --trim
```

Per-line trimming is the highest-value diagnostic for page-limit work: `agent-docx measure filing.md --lines --json` emits `deterministic.lines[]` with `page`, `ratio`, `unusedTwips`, `isLastLineOfBlock`. Filter client-side with `jq '[.deterministic.lines[] | select(.page==2)]'` or server-side via `--lines-page 2` (requires `--lines`); the human view shows a bar per line when not `--json`. `--paragraphs` / `--trim` are short-hand subsets of `--lines` (last-line only / ranked opportunities).

**Amortize iterative preflight — do not spawn per variant.** `measureMarkdown` caches Liberation Serif parsing (first ~90ms, ~30ms each after) inside one Node process. In an agent harness, import once and measure many variants without re-spawning:

```js
import { measureMarkdown } from "agent-docx";
const opts = { profile: "us-district-conventional", lines: true };
for (const md of variants) console.log(await measureMarkdown(md, opts)); // 15× ~600ms total
```

If you must use the CLI, batch in one spawn `printf '%s\n' '{"markdown":"# Hello"}' | agent-docx measure --batch --input-jsonl` or stream `agent-docx measure --watch filing.md --lines --jsonl`. See `skills/agent-docx/SKILL.md` (Performance section) for ext4 vs 9p benchmarks: `5× spawn 3.1s` vs `batch 0.69s`, library `15× 475ms`.
Portable estimates use pinned Liberation Serif 2.1.5 bytes as metrics while reporting the requested `Times New Roman` family and explicit substitution. Provide legally obtained custom font files through project configuration when a different deterministic metric source is required.

Deterministic pagination runs entirely in process:

1. Markdown is normalized into source-mapped legal blocks and footnotes.
2. Profile, supported template input, and configuration determine geometry, indentation, styles, line spacing, and page rules.
3. `fontkit` shapes runs with pinned font bytes; Unicode line breaking wraps text to usable width.
4. The paginator places lines, table rows, paragraph spacing, keeps, widow/orphan controls, explicit breaks, and counted-line caps.
5. Footnotes reserve bottom-page space when first referenced and report a relaxed constraint if an intrinsic split is unavoidable.
6. Results report physical pages, fractional equivalent-page use, visual and counted lines, paragraph-tail diagnostics, section attribution, and last-page metrics.

The same Markdown, configuration, and metric-font bytes produce the same deterministic result. The model reports unsupported content instead of guessing.

The `pnpm run accuracy` gate compares measured renderer page counts with each case's `targetPages` in the checked-in accuracy manifests. For `word` or `libreoffice`, an unavailable or errored requested renderer fails the run; the script never substitutes the deterministic or outer `pageCount` as a renderer measurement. Exact-match rate, mean absolute error, and worst-page error are computed against those manifest targets, while deterministic-only metrics are reported separately. Live Office gates require `AGENT_DOCX_TEST_WORD=1` or `AGENT_DOCX_TEST_LIBREOFFICE=1`; explicit test paths can be supplied with `AGENT_DOCX_ACCURACY_WORD_PATH` or `AGENT_DOCX_ACCURACY_LIBREOFFICE_PATH`.

## DOCX templates and import

`template inspect` uses a bounded ZIP/XML reader to inspect supported style inheritance, numbering, theme/font information, sections, header/footer relationships, fields, and caption components. Unsafe or unsupported package features—including macros, external relationships, embedded objects, scripts, encrypted parts, and arbitrary package copying—are not executed or copied into output.

Supported inspected layout/style data can be consumed by project configuration. A template is input for supported style and geometry semantics, not a host document to merge into an output package.

## PDF export and page verification

```sh
agent-docx export \
  --project agent-docx.json --document motion --revision HEAD \
  --mode pdf --output motion.pdf
```

`export --mode pdf` (agent action `docx.export` with `mode: "pdf"`) compiles the clean DOCX through the same strict semantic re-import gate, then renders it to PDF with the hardened headless LibreOffice invocation (isolated user profile, fixed locale and timezone, absolute executable path). The deterministic measurement remains the stable baseline: the result reports both page counts and their delta, and the PDF artifact is published through the same crash-recoverable staged store as DOCX exports, with `pdfSha256`, page count, and renderer provenance in the artifact provenance record. LibreOffice is an explicit local dependency for this mode—missing or failing renderers surface as `LIBREOFFICE_NOT_FOUND`/`LIBREOFFICE_RENDER_FAILED`/`LIBREOFFICE_TIMEOUT`, never a silent fallback. Office-rendered pagination is verification evidence, not a portable golden.

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
- `agent-docx/redline-import-result.schema.json`
- `agent-docx/filing-set.schema.json`
- `agent-docx/filing-set-validation.schema.json`
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
pnpm typecheck
pnpm format
pnpm test
pnpm accuracy
pnpm smoke:agent
pnpm verify:pack
pnpm social:preview
```

`pnpm social:preview` deterministically regenerates `docs/assets/agent-docx-social-preview.png`. Launch copy, release template, social posts, repository metadata, social-preview alt text, and upload instructions live in [`docs/marketing-kit.md`](docs/marketing-kit.md).

## License

MIT. See [LICENSE](LICENSE). Third-party notices are in [THIRD_PARTY_NOTICES.txt](THIRD_PARTY_NOTICES.txt).
