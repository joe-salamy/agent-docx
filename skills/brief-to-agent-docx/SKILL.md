---
name: brief-to-agent-docx
description: Draft a litigation brief in Markdown and compile to DOCX via agent-docx with measure → validate → export and optional redline/PDF.
version: 0.1.0
---

# brief-to-agent-docx

Thin alias over the `agent-docx` skill for the common "brief/motion → DOCX" path. For full protocol, profiles, rule packs, and MCP details, see `skills/agent-docx/SKILL.md`.

## When to use

A user wants to draft, check page/line limits, validate against a court rule pack, and export a filing-ready DOCX (or redline/PDF) from Markdown.

## Minimal sequence

```sh
pnpm add agent-docx
agent-docx skills install   # optional: copy skills/agent-docx into .omp/skills

cat > metadata.json <<'JSON'
{
  "court": "U.S. District Court for the Northern District of California",
  "jurisdiction": "US",
  "caseName": "Smith v. Jones",
  "docketNumber": "3:26-cv-00001",
  "documentTitle": "Motion to Dismiss",
  "filingDate": "2026-08-26",
  "parties": [], "counsel": [], "certificates": []
}
JSON

agent-docx project init --project agent-docx.json --document motion \
  --source motion.md --profile cand-civil --filing-kind motion-document \
  --rule-pack cand-civil@2026-05-01 --metadata metadata.json

agent-docx revision checkpoint --project agent-docx.json --document motion --base HEAD --author "Drafter" --message "init"
agent-docx draft guidance --project agent-docx.json --document motion
agent-docx validate --project agent-docx.json --document motion
agent-docx export --project agent-docx.json --document motion --revision HEAD --mode clean --output motion.docx
```

Measure-only preflight without a project:

```sh
agent-docx measure motion.md --profile cand-civil --paragraphs --sections --trim --json
```

Redline (two checkpoints, then `export --base <earlier> --mode redline`); PDF (`--mode pdf` requires local LibreOffice). See `agent-docx` skill for failure codes, caps, and `--renderer compare` opt-in.
