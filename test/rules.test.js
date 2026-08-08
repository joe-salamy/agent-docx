import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import {
  builtInProfiles,
  builtInRulePacks,
  compileMarkdown,
} from "../dist/index.js";
import { metadata } from "./helpers.js";

const metadataWithCounsel = {
  ...metadata,
  counsel: [{ id: "counsel", name: "A. Counsel" }],
};

test("rule packs bind validation to checked-in legal source snapshots", async () => {
  const pack = builtInRulePacks["cand-civil@2026-05-01"];
  const source = await readFile(
    new URL(`../assets/rules/${pack.sourceExcerpt}`, import.meta.url),
  );
  const sourceHash = `sha256:${(await import("node:crypto")).createHash("sha256").update(source).digest("hex")}`;
  assert.equal(sourceHash, pack.sourceSha256);

  const compiled = await compileMarkdown("# Motion\n\nRequested relief.\n", {
    documentId: "motion",
    profile: "cand-civil",
    filingKind: "motion-document",
    rulePack: "cand-civil@2026-05-01",
    metadata: metadataWithCounsel,
  });
  const footer = compiled.validation.findings.find(
    (finding) => finding.checkId === "cand.footer",
  );
  assert.equal(footer?.status, "fail");
  assert.equal(compiled.validation.status, "fail");
  assert.equal(
    compiled.validation.scope.sourceSnapshots[0]?.sha256,
    pack.sourceSha256,
  );
});

test("FRAP selects the monospaced predicate from verified pitch evidence", async () => {
  const profile = {
    ...builtInProfiles["frap-32"],
    id: "frap-mono-test",
    requestedFontFamily: "Courier New",
    maxCharactersPerInch: 11,
  };
  const compiled = await compileMarkdown("Brief text.\n", {
    documentId: "brief",
    profile,
    filingKind: "principal-brief",
    rulePack: "frap-32@2024-12-01",
    metadata: {
      ...metadata,
      certificates: [
        {
          id: "compliance",
          kind: "compliance",
          basis: "words",
          signerCounselId: "counsel",
        },
      ],
    },
  });
  const typeface = compiled.validation.findings.find(
    (finding) => finding.checkId === "frap32.typeface.monospaced",
  );
  assert.equal(typeface?.status, "fail");
  assert.equal(
    compiled.validation.findings.some(
      (finding) => finding.checkId === "frap32.typeface.proportional",
    ),
    false,
  );
});

test("cand.lines counts excluded blockquote lines out of the per-page cap", async () => {
  const single = { rule: "auto", numerator: 240, denominator: 240 };
  const base = builtInProfiles["cand-civil"];
  const profile = {
    ...base,
    id: "cand-counted-lines-test",
    body: { ...base.body, lineSpacing: single },
    blockquote: { ...base.blockquote, lineSpacing: single },
    list: { ...base.list, lineSpacing: single },
    footnote: { ...base.footnote, lineSpacing: single },
  };
  const body = Array.from({ length: 26 }, (_, index) => `Body line ${index}.`);
  const quotes = Array.from({ length: 4 }, (_, index) => `> Quote ${index}.`);
  const compiled = await compileMarkdown([...body, ...quotes, ""].join("\n"), {
    documentId: "motion",
    profile,
    filingKind: "motion-document",
    rulePack: "cand-civil@2026-05-01",
    metadata,
  });
  const lines = compiled.validation.findings.find(
    (finding) => finding.checkId === "cand.lines",
  );
  assert.equal(lines?.status, "pass");
  assert.equal(lines?.evidence.lines?.length, 1);
  assert.equal(lines?.evidence.lines[0], 26);
  assert.equal(
    lines?.evidence.countedLineSource,
    "deterministic-counted-lines",
  );
  assert.equal(lines?.evidence.maximum, 28);
});
