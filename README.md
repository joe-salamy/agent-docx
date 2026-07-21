# agent-docx

Deterministic DOCX-equivalent pagination for legal-prose Markdown, with opt-in Microsoft Word and LibreOffice Writer measurements.

```sh
pnpm add agent-docx
agent-docx filing.md
agent-docx filing.md --profile frap-32 --paragraphs --trim
```

```ts
import { estimateMarkdown, measureMarkdown, inspectDocxTemplate } from "agent-docx";
const estimate = await estimateMarkdown(markdown, { profile: "cand-civil" });
```

The default `us-district-conventional` profile is a product baseline: U.S. Letter, one-inch margins, 12-point Times-compatible serif metrics, and double spacing. It does not certify filing compliance. Verify current court, local, judge, document-type, and case-specific rules. `frap-32` and `cand-civil` carry source citations and effective dates.

Portable estimates use pinned Liberation Serif 2.1.5 bytes as metrics while reporting `Times New Roman` as the requested family and the substitution explicitly. Supply legally obtained font bytes for a different deterministic metric source. Word and Writer results are renderer-specific and include environment provenance; Writer is not treated as Word.

Use `--json` for one machine result, `--batch` for ordered JSONL, `--watch` for change streams, and `--inspect-template FILE.docx --json` for bounded read-only template inspection. Configuration is validated against the exported `agent-docx/config.schema.json` schema. No parent/home config or environment configuration is discovered.
