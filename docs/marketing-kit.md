# agent-docx launch kit

## Repository metadata

- **About:** Agent-first legal Markdown-to-DOCX compiler with deterministic page and line constraints, native redlines, embedded revisions, and optional Word validation.
- **Homepage:** https://www.npmjs.com/package/agent-docx
- **Topics:** `legal-tech`, `legal-drafting`, `markdown`, `docx`, `wordprocessingml`, `redline`, `tracked-changes`, `pagination`, `ai-agents`, `typescript`, `microsoft-word`, `libreoffice`
- **Social-preview heading:** `Draft legal DOCX with agents that know the page.`
- **Social-preview subline:** `Markdown → measured legal DOCX → native redline`
- **Social-preview alt text:** Dark navy agent-docx social preview with a blue vertical rule, the white agent-docx wordmark, the heading “Draft legal DOCX with agents that know the page.”, and the subline “Markdown → measured legal DOCX → native redline”.
- **Editable source:** `docs/assets/agent-docx-social-preview.svg`
- **Regenerate social preview:** `pnpm social:preview` (uses packaged Liberation Serif fonts and verifies a 1280×640 PNG signature below 1 MiB).

## Positioning

agent-docx is a local-first legal Markdown-to-DOCX compiler for agents, combining embedded revisions, native Word redlines, deterministic page and line intelligence, filing-rule checks, and optional Office validation in one auditable workflow.

### Target users

- Agent developers building controlled drafting, review, and filing workflows.
- Litigation teams that keep source in Markdown but deliver DOCX.
- Legal operations teams needing reproducible local artifacts, revision records, and machine-readable diagnostics.
- Lawyers using optional Word or LibreOffice checks without making either application a generation dependency.

## Six-step workflow

1. Draft source-mapped legal Markdown in an `agent-docx` project.
2. Evaluate proposed `SourcePatch` edits before changing the draft.
3. Apply a verified revision with a named author and message.
4. Validate the selected version against its pinned rule pack and structure checks.
5. Export a clean DOCX or a native-redline DOCX from immutable revisions.
6. Optionally verify the generated DOCX with locally installed Microsoft Word or LibreOffice.

## Differentiators

- Generates DOCX itself; Pandoc is not required.
- Uses deterministic, pinned font metrics and a source-mapped paginator by default.
- Stores immutable revision records, source/config/dependency snapshots, semantic change sets, and artifact provenance locally.
- Exports native Word tracked changes and review comments for revision comparisons.
- Separates rule-pack findings from layout profiles and binds implemented checks to checked-in source excerpts.
- Offers Word and LibreOffice only as explicit local validation engines.

## Local and security model

Projects use an on-disk manifest and a content-addressed `.agent-docx` store. Source, template, custom-font, and asset paths are restricted to regular non-symlink files inside the project. DOCX inspection and import use bounded ZIP/XML handling. The compiler does not fetch document assets or invoke a shell to generate DOCX. Word and LibreOffice are opt-in local processes; their versions, fonts, and results remain provenance rather than a hidden default.

## Supported claims

agent-docx supports source-mapped Markdown paragraphs, headings, blockquotes, ordered and unordered lists, GFM tables, thematic breaks, hard breaks, emphasis, strong emphasis, strikethrough, inline code, safe absolute links, footnotes, and controlled legal directives for captions, TOCs, TOAs, signatures, certificates, page and section breaks, numbered paragraphs, exhibits, length exclusions, images, references, and authorities.

It intentionally rejects arbitrary HTML, YAML front matter, code blocks, math, unsafe or relative ordinary links, unknown directives or attributes, unsupported DOCX constructs, and remote asset fetching. Filing-rule results are not legal advice or filing certification.

Native redline export supports source-mapped body paragraphs, headings, blockquotes, and controlled numbered paragraphs without footnotes. It deliberately rejects unsupported redline structures rather than producing a partial or misleading tracked-change file.
DOCX import is strict and fidelity-reported: recognized material is preserved, normalized, or externalized; unsupported structures are reported rather than silently dropped. External exhibits use hash-inventoried attachment bundles and are accepted only with an explicit matching bundle.

## Install and quick start

Requires Node.js 24 or newer.

```sh
pnpm add agent-docx

agent-docx project init \
  --document motion \
  --source motion.md \
  --profile cand-civil \
  --filing-kind motion-document \
  --rule-pack cand-civil@2026-05-01 \
  --metadata metadata.json

agent-docx revision checkpoint \
  --document motion \
  --base HEAD \
  --author "Drafter" \
  --message "Initial motion"

agent-docx draft guidance --document motion
agent-docx validate --document motion
agent-docx export --document motion --revision HEAD --mode clean --output motion.docx
```

For stateful agents, send one strict request per line:

```sh
printf '%s\n' \
  '{"schemaVersion":1,"id":"measure-1","action":"document.measure","project":"agent-docx.json","params":{"documentId":"motion"}}' \
  | agent-docx agent --input-jsonl
```

Published Draft 2020-12 schemas include the project, rule pack, revision, change-set, source-patch, validation, DOCX artifact/import, and agent request/response/watch contracts. Resolve them from the `agent-docx/<schema-name>.schema.json` package exports.

## Launch checklist

- [ ] Confirm `package.json` contains the About, homepage, and topic values above.
- [ ] Run `pnpm test` and `pnpm verify:pack`.
- [ ] Run an end-to-end project workflow: checkpoint, evaluate, validate, clean export, and redline export.
- [ ] Run optional Word or LibreOffice validation only on a host with the corresponding local application.
- [ ] Generate `docs/assets/agent-docx-social-preview.png` with `pnpm social:preview`.
- [ ] Upload that PNG through GitHub **General → Social preview → Edit**.
- [ ] Apply repository name, About, homepage, and topic metadata; verify the HTTPS remote.

## Release title and body template

**Title**

`agent-docx 0.1.0 — local-first legal Markdown-to-DOCX for agents`

**Body**

> agent-docx is a local-first legal Markdown-to-DOCX compiler for agents, combining embedded revisions, native Word redlines, deterministic page and line intelligence, filing-rule checks, and optional Office validation in one auditable workflow.
>
> It generates DOCX itself; Pandoc, Word, and LibreOffice are not required for generation. The default path is local and deterministic. Microsoft Word and LibreOffice remain explicit, optional validation engines because native pagination can vary by application, font installation, build, and printer. Filing-rule findings include their evidence and provenance, but they are not legal advice or filing certification.

## Social posts

### Short post 1

Agents can draft legal Markdown. The hard part is delivering a real Word file that fits. agent-docx compiles Markdown to DOCX itself, measures page and line constraints, and exports clean or native-redline files—local-first, no Pandoc. https://github.com/joe-salamy/agent-docx

### Short post 2

Markdown revisions should become real Word revisions. agent-docx turns semantic Markdown diffs into native tracked changes with author/date provenance, comments, and a machine-readable change set. The source stays Markdown; the deliverable is DOCX. https://github.com/joe-salamy/agent-docx

### Short post 3

Stop discovering a brief is over the limit after opening Word. agent-docx lets agents evaluate proposed edits against page, section, paragraph, and last-line budgets before changing the draft. https://github.com/joe-salamy/agent-docx

### Short post 4

agent-docx is deterministic by default and local-first. Word and LibreOffice are optional validators—not hidden dependencies and not the DOCX generator. Drafts, revision history, and artifacts stay on your machine. https://github.com/joe-salamy/agent-docx

## Long-post copy

### Opening

Today I’m releasing agent-docx, an agent-first legal Markdown-to-DOCX compiler. It gives drafting agents a source-mapped, page- and line-aware editing loop, embedded immutable revisions, semantic diffs, native Word redlines, legal document structure, and strict fidelity reports for DOCX import. It produces the DOCX itself; Pandoc, Word, and LibreOffice are not required for generation.

### Closing

The default path is local and deterministic. Microsoft Word and LibreOffice remain explicit, optional validation engines because native pagination can vary by application, font installation, build, and printer. Filing-rule findings include their evidence and provenance, but they are not legal advice or filing certification. https://github.com/joe-salamy/agent-docx
