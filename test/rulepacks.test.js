import assert from "node:assert/strict";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";
import { compileMarkdown, createProject } from "../dist/index.js";
import { validateUserRulePack } from "../dist/legal/rules.js";
import { metadata } from "./helpers.js";

const sourceHash = `sha256:${"0".repeat(64)}`;

const packWith = (checks) => ({
  id: "local-rules@2026-01-01",
  sourceUrl: "https://example.test/local-rules",
  effectiveDate: "2026-01-01",
  sourceSha256: sourceHash,
  sourceExcerpt: "local-rules.txt",
  checks,
  unmodeledProvisions: [],
});

const check = (id, kind, params) => ({
  id,
  kind,
  citation: `Local Rule ${id}`,
  predicate: `predicate for ${id}`,
  params,
});

const metadataAndSource = async (directory) => {
  await writeFile(join(directory, "motion.md"), "# Motion\n\nBody text.\n");
};

test("validateUserRulePack validates data-driven checks and rejects malformed packs", () => {
  const pack = packWith([
    check("size", "page-size", { widthTwips: 12240, heightTwips: 15840 }),
    check("lines", "counted-lines-maximum", { perPageMaximum: 28 }),
    check("footer", "required-footer", { requiredTokens: ["CASE"] }),
    check("block", "required-block", { kinds: ["paragraph"] }),
  ]);
  assert.equal(validateUserRulePack(pack, "valid").checks.length, 4);

  assert.throws(
    () =>
      validateUserRulePack(
        packWith([check("bad", "not-a-kind", {})]),
        "bad-kind",
      ),
    (error) => error?.code === "RULE_PACK_INVALID",
  );
  assert.throws(
    () =>
      validateUserRulePack(
        packWith([check("bad", "page-size", { widthTwips: 12240 })]),
        "bad-params",
      ),
    (error) => error?.code === "RULE_PACK_INVALID",
  );
  assert.throws(
    () =>
      validateUserRulePack(
        packWith([
          check("duplicate", "reference-integrity", {}),
          check("duplicate", "reference-integrity", {}),
        ]),
        "duplicate-ids",
      ),
    (error) => error?.code === "RULE_PACK_INVALID",
  );
});

test("project snapshots and executes custom rule packs", async () => {
  const directory = await mkdtemp(join(tmpdir(), "agent-docx-rulepacks-"));
  try {
    await metadataAndSource(directory);
    const passingPack = packWith([
      check("size", "page-size", { widthTwips: 12240, heightTwips: 15840 }),
      check("lines", "counted-lines-maximum", { perPageMaximum: 1000 }),
      check("footer", "required-footer", { requiredTokens: ["CASE"] }),
      check("block", "required-block", { kinds: ["paragraph"] }),
    ]);
    const failingPack = packWith([
      check("size", "page-size", { widthTwips: 1, heightTwips: 1 }),
      check("lines", "counted-lines-maximum", { perPageMaximum: 1 }),
      check("footer", "required-footer", { requiredTokens: ["MISSING"] }),
      check("block", "required-block", { kinds: ["signature"] }),
    ]);
    await writeFile(
      join(directory, "pass-pack.json"),
      JSON.stringify(passingPack),
    );
    await writeFile(
      join(directory, "fail-pack.json"),
      JSON.stringify(failingPack),
    );
    const project = await createProject(join(directory, "project.json"), {
      documentId: "motion",
      source: "motion.md",
      profile: "us-district-conventional",
      metadata,
      chrome: { footers: { default: "CASE" } },
      rulePacks: ["pass-pack.json"],
    });
    const passing = await project.validate("motion");
    assert.deepEqual(
      Object.fromEntries(
        passing.findings
          .filter((finding) =>
            ["size", "lines", "footer", "block"].includes(finding.checkId),
          )
          .map((finding) => [finding.checkId, finding.status]),
      ),
      { size: "pass", lines: "pass", footer: "pass", block: "pass" },
    );

    await project.configureDocument("motion", {
      baseRevision: null,
      changes: { rulePacks: ["fail-pack.json"] },
      author: { name: "Tester" },
      message: "Use failing local rules",
    });
    const failing = await project.validate("motion");
    assert.deepEqual(
      Object.fromEntries(
        failing.findings
          .filter((finding) =>
            ["size", "lines", "footer", "block"].includes(finding.checkId),
          )
          .map((finding) => [finding.checkId, finding.status]),
      ),
      { size: "fail", lines: "fail", footer: "fail", block: "fail" },
    );
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});

test("malformed and out-of-project rule packs are rejected", async () => {
  const directory = await mkdtemp(
    join(tmpdir(), "agent-docx-rulepacks-invalid-"),
  );
  try {
    await metadataAndSource(directory);
    await writeFile(
      join(directory, "malformed.json"),
      JSON.stringify({ nope: true }),
    );
    const malformedProject = await createProject(
      join(directory, "malformed-project.json"),
      {
        documentId: "motion",
        source: "motion.md",
        profile: "us-district-conventional",
        metadata,
        rulePacks: ["malformed.json"],
      },
    );
    await assert.rejects(
      () => malformedProject.validate("motion"),
      (error) => error?.code === "RULE_PACK_INVALID",
    );
    await assert.rejects(
      () =>
        createProject(join(directory, "outside-project.json"), {
          documentId: "motion",
          source: "motion.md",
          profile: "us-district-conventional",
          metadata,
          rulePacks: ["../outside.json"],
        }),
      (error) => error?.code === "PATH_OUTSIDE_PROJECT",
    );
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});

test("changing a checkpointed rule pack is a project hash conflict", async () => {
  const directory = await mkdtemp(join(tmpdir(), "agent-docx-rulepacks-hash-"));
  try {
    await metadataAndSource(directory);
    const pack = packWith([
      check("size", "page-size", { widthTwips: 12240, heightTwips: 15840 }),
    ]);
    const packPath = join(directory, "pack.json");
    await writeFile(packPath, JSON.stringify(pack));
    const project = await createProject(join(directory, "project.json"), {
      documentId: "motion",
      source: "motion.md",
      profile: "us-district-conventional",
      metadata,
      rulePacks: ["pack.json"],
    });
    await project.checkpoint("motion", {
      baseRevision: null,
      author: { name: "Tester" },
      message: "Checkpoint local rules",
    });
    await writeFile(
      packPath,
      JSON.stringify({ ...pack, sourceExcerpt: "changed.txt" }),
    );
    await assert.rejects(
      () => project.validate("motion"),
      (error) =>
        error?.code === "PROJECT_INVALID" &&
        /rule pack changed since snapshot/i.test(error.message),
    );
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});

test("built-in FRAP rule pack validation remains available", async () => {
  const compiled = await compileMarkdown("Brief text.\n", {
    documentId: "brief",
    profile: "frap-32",
    filingKind: "principal-brief",
    rulePack: "frap-32@2024-12-01",
    metadata,
  });
  assert.equal(compiled.validation.rulePack, "frap-32@2024-12-01");
  assert.ok(
    compiled.validation.findings.some(
      (finding) => finding.checkId === "frap32.page-size",
    ),
  );
});
