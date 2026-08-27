---
name: agent-docx
description: Agent-first legal Markdown to DOCX with deterministic pagination, revision-bound drafting, rule-pack validation, and native redlines. Use when drafting, measuring, validating, or exporting litigation documents.
version: 0.1.0
---

# agent-docx

Use when an agent or human needs to draft in legal Markdown and produce deterministic DOCX — page/line preflight, court rule-pack validation, revision provenance, semantic diffs, native tracked changes, or PDF verification.

## Install

Requires Node 24+.

```sh
pnpm add agent-docx
agent-docx --version
agent-docx --help  # must be used alone; subcommand help is not supported
```

The npm package is the canonical distribution. Working from a checkout, invoke `node dist/cli.js <command>` directly — no install step needed.

### Optional: install this skill into the consumer repo

The package ships this skill under `skills/agent-docx/SKILL.md`. Copy it where your harness discovers skills (opt-in, no postinstall side effect):

```sh
# opencode / omp — project-local (recommended)
agent-docx skills install
agent-docx skills install --dest .omp/skills --force

# Claude Code
agent-docx skills install --dest .claude/skills

# global
agent-docx skills install --global

# preview without writing
agent-docx skills install --dry-run --json
```

`install` copies the versioned skill from the installed package; `list` shows available skills. Re-run after `pnpm update agent-docx` to pick up skill updates. The skill files are part of the `agent-docx` package `files` — they are not generated and never overwrite without `--force`.

## Core workflow (revision-bound)

Every mutation is optimistic and revision-bound: you must name the expected `--base` revision; a stale base is rejected rather than silently overwritten. The project manifest (`agent-docx.json`) owns configuration; `.agent-docx/` is a private content-addressed object store — do not edit it by hand.

```sh
# 1. metadata.json — five nonempty text fields required; arrays may be empty
cat > metadata.json <<'JSON'
{
  "court": "U.S. District Court for the Northern District of California",
  "jurisdiction": "US",
  "caseName": "Smith v. Jones",
  "docketNumber": "3:26-cv-00001",
  "documentTitle": "Motion to Dismiss",
  "filingDate": "2026-08-26",
  "parties": [],
  "counsel": [],
  "certificates": []
}
JSON

# 2. init project + first document
agent-docx project init \
  --project agent-docx.json \
  --document motion \
  --source motion.md \
  --profile cand-civil \
  --filing-kind motion-document \
  --rule-pack cand-civil@2026-05-01 \
  --metadata metadata.json

# add a second document to the same project
agent-docx project add \
  --project agent-docx.json \
  --document brief \
  --source brief.md \
  --profile cand-civil \
  --filing-kind motion-document \
  --rule-pack cand-civil@2026-05-01 \
  --metadata metadata.json

# 3. checkpoint — creates immutable revision (content-hashed ID)
agent-docx revision checkpoint \
  --project agent-docx.json --document motion \
  --base HEAD --author "Drafter" --message "Initial motion"

# 4. draft loop — guidance before/after edits
agent-docx draft guidance --project agent-docx.json --document motion
# propose a patch, verify deterministically, then apply only if hash + base still match
agent-docx draft evaluate --project agent-docx.json --document motion --patch patch.json
agent-docx draft apply --project agent-docx.json --document motion --patch patch.json --expected-hash <hash> --base HEAD --author "Drafter" --message "Tighten argument"

# 5. validate against pinned rule pack(s)
agent-docx validate --project agent-docx.json --document motion

# 6. export — clean, redline, or pdf (pdf requires local LibreOffice)
agent-docx export --project agent-docx.json --document motion --revision HEAD --mode clean --output motion.docx
agent-docx export --project agent-docx.json --document motion --revision HEAD --base <earlier-rev> --mode redline --output motion-redline.docx
agent-docx export --project agent-docx.json --document motion --revision HEAD --mode pdf --output motion.pdf

# 7. filing sets (optional ordered group with shared page cap)
agent-docx filing-set add --project agent-docx.json --id motion-package --label "Motion package" --documents motion,brief --page-cap 30
agent-docx filing-set validate --project agent-docx.json --id motion-package
```

`document configure --base HEAD --changes changes.json` records config changes as a new revision. `revision diff BASE HEAD` is a semantic `ChangeSet` (add/remove/replace/move/config/dependency/annotation), not a line diff. `revision restore` creates a new revision restoring a target; `revision resolve` applies accept/reject decisions. `review add --block <id> --start 0 --end 18` anchors comments to legal blocks.

## Standalone measure (no project)

**Primary way to trim to a page limit — use per-line fill, not word count:**
```sh
# Full per-line table: page, ratio (used/available), twips slack, isLastLineOfBlock
agent-docx measure filing.md --profile cand-civil --lines --json | jq '[.deterministic.lines[] | {page,ratio,unusedTwips,isLastLineOfBlock,text}]'

# Only slackest last lines (best edits): sort by ratio then fix lowest 10
agent-docx measure filing.md --lines --json | jq '[.deterministic.lines[] | select(.isLastLineOfBlock)] | sort_by(.ratio) | .[0:10]'

# Scope to page 2 (or --lines-page 2 for server-side filter)
agent-docx measure filing.md --lines --lines-page 2 --json | jq .
agent-docx measure filing.md --lines --json | jq '[.deterministic.lines[] | select(.page==2)]'

# Human bar view
agent-docx measure filing.md --profile cand-civil --lines
```
`--paragraphs` / `--trim` are short-hand subsets of `--lines` (last-line only / ranked opportunities). `--sections` remains separate for section page breakdown.

Agents MUST prefer `deterministic.lines[].ratio` + `unusedTwips` + `estimatedRemovalTwips` (via `paragraphs.oneLineReduction`) over `wordCount/wordsPerPage`. A 25% last line wastes ~0.75*availableTwips; deleting that many twips of text collapses one visual line.

```sh
agent-docx measure filing.md --profile cand-civil --paragraphs --sections --trim --json
agent-docx measure filing.md --profile us-district-conventional --renderer compare --json
printf '%s\n' '{"path":"filing.md"}' | agent-docx measure --batch --input-jsonl
agent-docx measure --batch "briefs/*.md" --include "*.md" --exclude "draft/*"
echo "# Hello" | agent-docx measure - --profile us-district-conventional --json
```

Stateless JSONL batch and watch streaming are also available: `agent-docx agent --input-jsonl` (one closed JSON request/response per line) and `agent-docx agent --watch --project FILE --document ID --jsonl`.

## Profiles, rule packs, and validation

- **Profiles** control geometry/styles: `us-district-conventional` (baseline, Letter, 1in, 12pt, double), `cand-civil`, `frap-32`. `agent-docx profiles --json` lists catalog.
- **Rule packs** are versioned snapshots (`frap-32@2024-12-01`, `cand-civil@2026-05-01`) with URL, effective date, source excerpt, SHA-256, modeled predicates, and unmodeled provisions. Custom packs are schema-validated JSON (`rule-pack.schema.json`) attached via `document configure --changes '{"rulePacks":["packs/firm.json"]}'` — content-hashed into revisions.
- **Validation** reports `ValidationResult` findings with `checkId`, `status: pass|fail|unknown`, `severity`, `source/evidence`, `remediation`, and `scope.sourceSnapshots`. `unknown` means the deterministic compiler cannot establish native behavior — not a pass. Pin rule-pack versions; a changing court website never silently changes semantics.

## Markdown that compiles

Supported: paragraphs, headings, blockquotes, ordered/unordered lists, GFM tables, thematic breaks, hard breaks, emphasis/strong/strikethrough/inline code, safe absolute `https://` links, footnotes, and directives: `caption`, `toc`, `toa`, `pagebreak`, `signature` (known counsel ID), `certificate` (known cert ID), `sectionbreak` (`kind="next-page"|"continuous"`), `numbered` (seq level 1–4), `exhibit`/`image` (project-relative, image needs positive twip dims + alt), `length-exclusion` (closed set), inline legal references/authorities.

Rejected with source-aware `code` + `details.position`: arbitrary HTML, YAML front matter, fenced code blocks, math, relative/unsafe links, remote assets, unknown directives/attributes. Fix the source; do not loosen the parser.

## Deterministic pagination

In-process, no Office required: Markdown → source-mapped legal blocks → profile/template/config geometry → `fontkit` shaping with pinned Liberation Serif 2.1.5 (reports `Times New Roman`, substitution explicit) → paginator (lines, table rows, keeps, widow/orphan, breaks, counted-line caps, footnote reservation). Same Markdown+config+font bytes ⇒ same `pageCount`, `fractionalEquivalent`, diagnostics. `measureMarkdown({includeGeneratedDocx,sectionDiagnostics})` returns the same result as the CLI; `generatedDocx` bytes are only embedded when requested.

## DOCX, redline, and import

- `template inspect FILE.docx --json` inspects supported style/numbering/theme/sections safely (no macros/external relationships/embedded objects executed or copied; 512 entries / 25 MiB / 4 MiB per XML part caps).
- Redline export converts a resolved `ChangeSet` into native OOXML insertions/deletions + Word comments. Supports body paragraphs/headings/blockquotes/numbered paragraphs without footnotes; rejects lists/tables/images/exhibits/breaks/captions/TOC/TOA/footnotes with `DOCX_REDLINE_UNSUPPORTED` — use clean export instead.
- `import --inspect-only` is stateless; normal `import` requires target document/output/author/message. `import-redline` validates semantic manifest, requires HEAD + clean working copy, returns `ChangeSet` + `decisions` covering every change (byte-identical to `revision resolve` derivation).
- `export --mode pdf` re-imports the clean DOCX through the strict gate, then renders via headless LibreOffice (isolated profile, fixed locale/timezone, absolute path) — deterministic `pageCount` vs `pdfPageCount` delta reported.

## Optional Office validation

Default `--renderer deterministic` never launches Office. Opt in: `--renderer word|libreoffice|compare`. `--renderer compare` keeps deterministic top-level fields and adds `word`/`libreoffice` diagnostics; missing/errored requested renderer fails the run — the CLI never substitutes another source.

## Agent protocol and MCP

- **CLI JSONL agent transport**: `agent-docx agent --input-jsonl` — actions `project.init/add/get`, `document.configure/get/measure/validate`, `revision.checkpoint/list/get/restore/diff/resolve`, `draft.guidance/evaluate/apply`, `review.add/resolve`, `docx.export/import/inspect/importRedline`, `filingSet.add/remove/get/validate`. Envelope always has `schemaVersion,kind,sequence,requestId,action,project,documentId,revision`. Generated DOCX bytes never embedded — artifact `path` + `sha256` instead.
- **Watch**: `agent-docx agent --watch --project FILE --document ID --jsonl` streams `ready` (inventory), debounced `document.measure` results/errors, `end`.
- **MCP server**: `agent-docx mcp` — stdio JSON-RPC, every agent action becomes a tool (e.g. `document.validate`, `draft.evaluate`, `docx.export`, `filingSet.validate`). `tools/list` + `tools/call` with `params` + optional cwd-relative `project`; paths confined to server cwd; tool results as text + `structuredContent`; `isError:true` carries `{code,message}`.

## Failure handling

All failures are closed JSON on stderr, exit 1 (3 for `fail-over-limit` budget breach, 4 for Word/LibreOffice errors): `{"schemaVersion":1,"kind":"fatal","error":{"code":"...","message":"...","details":{...}}}`. Measure/batch errors are per-record `error` records; batch never partially processes a rejected input. Input caps: 64 MiB Markdown/stdin, 8 MiB per JSONL line; DOCX: 25 MiB compressed, 512 entries, 64 MiB decompressed, 4 MiB per XML part.

## Project library (for hosts that embed agent-docx)

```ts
import { createProject, openProject, measureMarkdown, compileMarkdown, estimateMarkdown } from "agent-docx";
import type { ProjectMeasureOptions, SerializableMeasurementResult } from "agent-docx";
```

`measureMarkdown(markdown, {profile, ...})` is the programmatic equivalent of `agent-docx measure`. Use `compileMarkdown` / `generateDocx` / `generateRedlineDocx` for direct DOCX bytes when you own the workflow outside the project store.

## References

- `agent-docx --help`, `README.md`, `schemas/*.schema.json` (Draft 2020-12), `profiles`, `rule-pack` and `template inspect` docs in this repo.
- Filing-rule findings are software evidence, not legal advice — confirm current court/local/judge/case requirements before filing.
