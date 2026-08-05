import assert from "node:assert/strict";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";
import { createProject, openProject } from "../dist/index.js";
import { metadata } from "./helpers.js";

const profile = "us-district-conventional";

const expectCode = async (operation, code) => {
  await assert.rejects(operation, (error) => {
    assert.equal(error?.code, code);
    return true;
  });
};

test("filing sets group ordered documents and enforce aggregate page caps", async () => {
  const directory = await mkdtemp(join(tmpdir(), "agent-docx-filing-sets-"));
  const manifestPath = join(directory, "agent-docx.json");
  try {
    await writeFile(join(directory, "motion.md"), "# Motion\n\nMotion body.\n");
    const project = await createProject(manifestPath, {
      documentId: "motion",
      source: "motion.md",
      profile,
      metadata,
    });
    await writeFile(
      join(directory, "declaration.md"),
      "# Declaration\n\nDeclaration body.\n",
    );
    await project.addDocument({
      documentId: "declaration",
      source: "declaration.md",
      profile,
      metadata: { ...metadata, documentTitle: "Declaration" },
    });

    const filingProject = await openProject(manifestPath);
    const firstState = await filingProject.addFilingSet({
      id: "primary",
      label: "Primary filing",
      documentIds: ["motion", "declaration"],
      pageCap: 1,
    });
    assert.deepEqual(firstState.filingSets, [
      {
        id: "primary",
        label: "Primary filing",
        documentIds: ["motion", "declaration"],
        pageCap: 1,
      },
    ]);
    const manifest = JSON.parse(await readFile(manifestPath, "utf8"));
    assert.deepEqual(manifest.filingSets, [
      {
        id: "primary",
        label: "Primary filing",
        documentIds: ["motion", "declaration"],
        pageCap: 1,
      },
    ]);

    const beforeCheckpoint = await filingProject.getFilingSet("primary");
    assert.deepEqual(beforeCheckpoint.documentIds, ["motion", "declaration"]);
    assert.equal(beforeCheckpoint.label, "Primary filing");
    assert.equal(beforeCheckpoint.pageCap, 1);
    assert.equal(beforeCheckpoint.documents.length, 2);
    for (const document of beforeCheckpoint.documents) {
      assert.equal(document.head, null);
      assert.match(document.workingTreeHash, /^sha256:[a-f0-9]{64}$/);
      assert.equal(document.matchesHead, false);
    }

    await expectCode(
      filingProject.addFilingSet({ id: "primary", documentIds: ["motion"] }),
      "PROJECT_INVALID",
    );
    await expectCode(
      filingProject.addFilingSet({ id: "unknown", documentIds: ["missing"] }),
      "PROJECT_INVALID",
    );
    await expectCode(
      filingProject.addFilingSet({
        id: "duplicate",
        documentIds: ["motion", "motion"],
      }),
      "PROJECT_INVALID",
    );
    await expectCode(
      filingProject.addFilingSet({ id: "Bad ID", documentIds: ["motion"] }),
      "INVALID_ARGUMENT",
    );

    await project.checkpoint("motion", {
      baseRevision: null,
      author: { name: "Drafter" },
      message: "Checkpoint motion",
    });
    await project.checkpoint("declaration", {
      baseRevision: null,
      author: { name: "Drafter" },
      message: "Checkpoint declaration",
    });

    const capped = await filingProject.validateFilingSet("primary");
    assert.equal(capped.pageCap?.limit, 1);
    assert.equal(capped.pageCap?.totalPages, 2);
    assert.equal(capped.pageCap?.status, "fail");
    assert.equal(capped.status, "fail");
    assert.equal(capped.documents.length, 2);
    assert.ok(
      capped.documents.every(
        (document) => document.validation?.status === "pass",
      ),
    );
    assert.ok(capped.documents.every((document) => document.pageCount === 1));

    const within = await filingProject.addFilingSet({
      id: "within",
      documentIds: ["declaration", "motion"],
      pageCap: 2,
    });
    assert.deepEqual(
      within.filingSets.map((set) => set.id),
      ["primary", "within"],
    );
    const withinValidation = await filingProject.validateFilingSet("within");
    assert.equal(withinValidation.pageCap?.totalPages, 2);
    assert.equal(withinValidation.pageCap?.status, "pass");
    assert.equal(withinValidation.status, "pass");

    await filingProject.addFilingSet({
      id: "uncapped",
      documentIds: ["motion", "declaration"],
    });
    const uncapped = await filingProject.validateFilingSet("uncapped");
    assert.equal(uncapped.pageCap, null);
    assert.equal(uncapped.status, "pass");

    const removed = await filingProject.removeFilingSet("primary");
    assert.deepEqual(
      removed.filingSets.map((set) => set.id),
      ["uncapped", "within"],
    );
    await expectCode(filingProject.getFilingSet("primary"), "PROJECT_INVALID");
    await expectCode(
      filingProject.removeFilingSet("missing"),
      "PROJECT_INVALID",
    );

    const reopened = await openProject(manifestPath);
    assert.deepEqual(
      (await reopened.getState()).filingSets.map((set) => set.id),
      ["uncapped", "within"],
    );
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});
