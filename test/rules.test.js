import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import { builtInProfiles, builtInRulePacks, compileMarkdown } from "../dist/index.js";

const metadata = {
  court: "United States District Court",
  jurisdiction: "Northern District of California",
  caseName: "Example v. Example",
  docketNumber: "3:26-cv-00001",
  documentTitle: "Motion",
  parties: [],
  counsel: [{ id: "counsel", name: "A. Counsel" }],
  certificates: [],
};

test("rule packs bind validation to checked-in legal source snapshots", async () => {
  const pack = builtInRulePacks["cand-civil@2026-05-01"];
  const source = await readFile(`assets/rules/${pack.sourceExcerpt}`);
  const sourceHash = `sha256:${(await import("node:crypto")).createHash("sha256").update(source).digest("hex")}`;
  assert.equal(sourceHash, pack.sourceSha256);

  const compiled = await compileMarkdown("# Motion\n\nRequested relief.\n", {
    documentId: "motion",
    profile: "cand-civil",
    filingKind: "motion-document",
    rulePack: "cand-civil@2026-05-01",
    metadata,
  });
  const footer = compiled.validation.findings.find((finding) => finding.checkId === "cand.footer");
  assert.equal(footer?.status, "fail");
  assert.equal(compiled.validation.status, "fail");
  assert.equal(compiled.validation.scope.sourceSnapshots[0]?.sha256, pack.sourceSha256);
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
      certificates: [{
        id: "compliance",
        kind: "compliance",
        basis: "words",
        signerCounselId: "counsel",
      }],
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
