import assert from "node:assert/strict";
import { mkdir, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";
import { createProject } from "../dist/index.js";
import { readDocxParts, repackDocxParts } from "../dist/docx/package.js";

const metadata = {
  court: "United States District Court",
  jurisdiction: "Northern District of California",
  caseName: "Example v. Example",
  docketNumber: "3:26-cv-00001",
  documentTitle: "Motion",
  parties: [],
  counsel: [],
  certificates: [],
};

const xmlFor = async (bytes) => {
  const parts = await readDocxParts(bytes);
  const xml = new TextDecoder().decode(parts.get("word/document.xml"));
  return { parts, xml };
};

const rewriteDocumentXml = async (bytes, rewrite) => {
  const { parts, xml } = await xmlFor(bytes);
  const next = rewrite(xml);
  assert.notEqual(next, xml);
  const updated = new Map(parts);
  updated.set("word/document.xml", new TextEncoder().encode(next));
  return repackDocxParts(updated);
};

const acceptedRedline = (bytes) =>
  rewriteDocumentXml(bytes, (xml) =>
    xml
      .replace(/<w:ins\b[^>]*>[\s\S]*?<\/w:ins>/g, (value) =>
        value.replace(/^<w:ins\b[^>]*>/, "").replace(/<\/w:ins>$/, ""),
      )
      .replace(/<w:del\b[^>]*>[\s\S]*?<\/w:del>/g, (value) =>
        value
          .replace(/^<w:del\b[^>]*>/, "")
          .replace(/<\/w:del>$/, "")
          .replace(/<w:delText\b[^>]*>[\s\S]*?<\/w:delText>/g, ""),
      ),
  );

const rejectedRedline = (bytes) =>
  rewriteDocumentXml(bytes, (xml) =>
    xml
      .replace(/<w:ins\b[^>]*>[\s\S]*?<\/w:ins>/g, "")
      .replace(/<w:del\b[^>]*>/g, "")
      .replace(/<\/w:del>/g, "")
      .replace(/<w:delText\b/g, "<w:t")
      .replace(/<\/w:delText>/g, "</w:t>"),
  );

const setup = async (directory, withComment = false) => {
  const manifestPath = join(directory, "agent-docx.json");
  const sourcePath = join(directory, "motion.md");
  await writeFile(
    sourcePath,
    "# Motion\n\nFirst old sentence.\n\nSecond sentence.\n",
  );
  const project = await createProject(manifestPath, {
    documentId: "motion",
    source: "motion.md",
    profile: "us-district-conventional",
    metadata,
  });
  const base = await project.checkpoint("motion", {
    baseRevision: null,
    author: { name: "Drafter" },
    message: "Initial draft",
  });
  let reviewBase = base;
  if (withComment) {
    const baseDocument = await project.getDocument("motion", base.revision.id);
    reviewBase = await project.addReview("motion", {
      revision: base.revision.id,
      blockId: baseDocument.document.blocks[1].id,
      author: { name: "Reviewer" },
      message: "Please confirm this sentence.",
    });
  }
  await writeFile(
    sourcePath,
    (await readFile(sourcePath, "utf8"))
      .replace("First old sentence.", "First new sentence.")
      .replace("Second sentence.", "Second sentence plus addition."),
  );
  const head = await project.checkpoint("motion", {
    baseRevision: reviewBase.revision.id,
    author: { name: "Drafter" },
    message: "Revise motion",
  });
  const baseSource = (await project.getDocument("motion", base.revision.id))
    .source;
  const headSource = (await project.getDocument("motion", head.revision.id))
    .source;
  const exported = await project.exportDocx("motion", {
    revision: head.revision.id,
    mode: "redline",
    baseRevision: base.revision.id,
    output: join(directory, "motion-redline.docx"),
  });
  return { project, base, head, baseSource, headSource, bytes: exported.bytes };
};

test("reviewer redline import extracts accepted and rejected decisions", async () => {
  const root = await mkdtemp(join(tmpdir(), "agent-docx-redline-roundtrip-"));
  try {
    for (const mode of ["accept", "reject"]) {
      const directory = join(root, mode);
      await mkdir(directory);
      const setupState = await setup(directory);
      const reviewedBytes =
        mode === "accept"
          ? await acceptedRedline(setupState.bytes)
          : await rejectedRedline(setupState.bytes);
      const imported = await setupState.project.importRedline({
        documentId: "motion",
        input: reviewedBytes,
        author: { name: "Importer" },
        message: `Import ${mode} review`,
      });
      assert.equal(imported.resolution, "complete");
      const expectedDecision = mode === "accept" ? "accept" : "reject";
      assert.ok(
        Object.values(imported.decisions).every(
          (decision) => decision === expectedDecision,
        ),
      );
      assert.deepEqual(
        imported.changeSet,
        await setupState.project.diff(
          "motion",
          setupState.base.revision.id,
          setupState.head.revision.id,
        ),
      );
      const resolved = await setupState.project.resolveChanges("motion", {
        changeSet: imported.changeSet,
        decisions: imported.decisions,
        author: { name: "Importer" },
        message: `Resolve ${mode} review`,
      });
      const finalSource = (
        await setupState.project.getDocument("motion", resolved.revision.id)
      ).source;
      assert.equal(
        finalSource,
        mode === "accept" ? setupState.headSource : setupState.baseSource,
      );
    }
  } finally {
    await rm(root, { recursive: true, force: true });
  }
});

test("untouched redline reports no resolution and preserves comments", async () => {
  const directory = await mkdtemp(join(tmpdir(), "agent-docx-redline-none-"));
  try {
    const setupState = await setup(directory, true);
    const before = await setupState.project.getDocument(
      "motion",
      setupState.head.revision.id,
    );
    const imported = await setupState.project.importRedline({
      documentId: "motion",
      input: setupState.bytes,
      author: { name: "Importer" },
      message: "Inspect untouched review",
    });
    assert.equal(imported.resolution, "none");
    assert.deepEqual(imported.decisions, {});
    assert.deepEqual(imported.annotations, before.annotations);
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});

test("foreign edits in resolved redlines are rejected as ambiguous", async () => {
  const directory = await mkdtemp(
    join(tmpdir(), "agent-docx-redline-ambiguous-"),
  );
  try {
    const setupState = await setup(directory);
    const foreign = await rewriteDocumentXml(setupState.bytes, (xml) =>
      xml.replace(/(<w:t[^>]*>)new(<\/w:t>)/, "$1Foreign$2"),
    );
    await assert.rejects(
      setupState.project.importRedline({
        documentId: "motion",
        input: await acceptedRedline(foreign),
        author: { name: "Importer" },
        message: "Import ambiguous review",
      }),
      (error) => error?.code === "DOCX_IMPORT_UNSUPPORTED",
    );
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});

test("redline import requires the semantic revision to remain HEAD", async () => {
  const directory = await mkdtemp(
    join(tmpdir(), "agent-docx-redline-conflict-"),
  );
  try {
    const setupState = await setup(directory);
    const sourcePath = join(directory, "motion.md");
    await writeFile(
      sourcePath,
      `${await readFile(sourcePath, "utf8")}\nA later edit.\n`,
    );
    await setupState.project.checkpoint("motion", {
      baseRevision: setupState.head.revision.id,
      author: { name: "Drafter" },
      message: "Move HEAD",
    });
    await assert.rejects(
      setupState.project.importRedline({
        documentId: "motion",
        input: setupState.bytes,
        author: { name: "Importer" },
        message: "Import stale review",
      }),
      (error) => error?.code === "REVISION_CONFLICT",
    );
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});
