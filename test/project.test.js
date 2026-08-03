import assert from "node:assert/strict";
import { mkdir, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";
import { compileMarkdown, createProject, openProject } from "../dist/index.js";
import { readDocxParts, repackDocxParts } from "../dist/docx/package.js";
import { inspectDocxMaterial } from "../dist/docx/import.js";

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

test("project checkpoints source-mapped legal documents and review revisions", async () => {
  const directory = await mkdtemp(join(tmpdir(), "agent-docx-project-"));
  const manifestPath = join(directory, "agent-docx.json");
  const sourcePath = join(directory, "motion.md");
  try {
    await writeFile(sourcePath, "# Motion\n\nThis is the body.\n");
    const project = await createProject(manifestPath, {
      documentId: "motion",
      source: "motion.md",
      profile: "us-district-conventional",
      metadata,
    });
    assert.equal((await project.getState()).documents[0].head, null);
    assert.match(await readFile(sourcePath, "utf8"), /agent-docx:block/);

    const first = await project.checkpoint("motion", {
      baseRevision: null,
      author: { name: "Drafter" },
      message: "Initial draft",
    });
    assert.match(first.revision.id, /^sha256:[a-f0-9]{64}$/);
    const historical = await project.getDocument("motion", first.revision.id);
    assert.equal(historical.document.blocks.length, 2);
    assert.equal(historical.revision, first.revision.id);
    assert.match(await readFile(sourcePath, "utf8"), /agent-docx:block/);

    const review = await project.addReview("motion", {
      revision: "HEAD",
      blockId: historical.document.blocks[1].id,
      range: { start: 0, end: 4 },
      author: { name: "Reviewer" },
      message: "Check the title",
    });
    const reopened = await openProject(manifestPath);
    const latest = await reopened.getDocument("motion", "HEAD");
    assert.equal(latest.head, review.revision.id);
    assert.equal(latest.annotations.length, 1);
    assert.equal(latest.annotations[0].status, "open");
    assert.match((await reopened.getRevision("motion", "HEAD")).deltaObject, /^sha256:[a-f0-9]{64}$/);
    assert.equal((await reopened.listRevisions("motion")).items.length, 2);
    const markedSource = await readFile(sourcePath, "utf8");
    await writeFile(sourcePath, markedSource.replace("This is the body.", "This is the revised body."));
    const edited = await reopened.checkpoint("motion", {
      baseRevision: review.revision.id,
      author: { name: "Drafter" },
      message: "Revise body",
    });
    assert.deepEqual((await reopened.getDocument("motion", "HEAD")).annotations[0].range, {
      start: 0,
      end: 4,
    });
    const changes = await reopened.diff("motion", review.revision.id, edited.revision.id);
    assert.ok(changes.changes.length > 0);
    const resolved = await reopened.resolveChanges("motion", {
      changeSet: changes,
      decisions: Object.fromEntries(
        [...changes.changes, ...changes.annotations].map((change) => [change.id, "reject"]),
      ),
      author: { name: "Reviewer" },
      message: "Reject revision",
    });
    assert.match(resolved.revision.resolutionObject, /^sha256:[a-f0-9]{64}$/);
    assert.match(await readFile(sourcePath, "utf8"), /This is the body\./);
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});
test("project diffs ignore source-offset shifts in unchanged blocks", async () => {
  const directory = await mkdtemp(join(tmpdir(), "agent-docx-diff-offsets-"));
  const manifestPath = join(directory, "agent-docx.json");
  const sourcePath = join(directory, "motion.md");
  try {
    await writeFile(
      sourcePath,
      "# Motion\n\nFirst statement.\n\nSecond statement.\n",
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
    const markedSource = await readFile(sourcePath, "utf8");
    await writeFile(
      sourcePath,
      markedSource.replace("First statement.", "A materially longer first statement."),
    );
    const head = await project.checkpoint("motion", {
      baseRevision: base.revision.id,
      author: { name: "Drafter" },
      message: "Revise first statement",
    });
    const changes = await project.diff("motion", base.revision.id, head.revision.id);
    assert.deepEqual(changes.changes.map((change) => change.kind), ["replace-text"]);
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});

test("project restore creates a new head from a reachable revision", async () => {
  const directory = await mkdtemp(join(tmpdir(), "agent-docx-restore-"));
  const manifestPath = join(directory, "agent-docx.json");
  const sourcePath = join(directory, "motion.md");
  try {
    await writeFile(sourcePath, "# Motion\n\nInitial body.\n");
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
    await writeFile(
      sourcePath,
      (await readFile(sourcePath, "utf8")).replace("Initial", "Revised"),
    );
    const head = await project.checkpoint("motion", {
      baseRevision: base.revision.id,
      author: { name: "Drafter" },
      message: "Revise body",
    });
    const restored = await project.restore("motion", {
      baseRevision: head.revision.id,
      targetRevision: base.revision.id,
      author: { name: "Reviewer" },
      message: "Restore initial draft",
    });
    assert.deepEqual(restored.revision.parents, [head.revision.id]);
    assert.match(await readFile(sourcePath, "utf8"), /Initial body\./);
    assert.equal((await project.getDocument("motion", "HEAD")).revision, restored.revision.id);
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});

test("project add owns a distinct source and can change the default document", async () => {
  const directory = await mkdtemp(join(tmpdir(), "agent-docx-add-"));
  const manifestPath = join(directory, "agent-docx.json");
  const motionPath = join(directory, "motion.md");
  const replyPath = join(directory, "reply.md");
  try {
    await writeFile(motionPath, "# Motion\n\nOpening body.\n");
    await writeFile(replyPath, "# Reply\n\nReply body.\n");
    const project = await createProject(manifestPath, {
      documentId: "motion",
      source: "motion.md",
      profile: "us-district-conventional",
      metadata,
    });
    const state = await project.addDocument({
      documentId: "reply",
      source: "reply.md",
      profile: "us-district-conventional",
      filingKind: "reply-brief",
      metadata: { ...metadata, documentTitle: "Reply" },
      makeDefault: true,
    });
    assert.equal(state.manifest.defaultDocument, "reply");
    assert.deepEqual(
      state.documents.map((document) => document.documentId),
      ["motion", "reply"],
    );
    assert.match(await readFile(replyPath, "utf8"), /agent-docx:block/);
    await assert.rejects(
      project.addDocument({
        documentId: "reply",
        source: "reply.md",
        profile: "us-district-conventional",
        metadata: { ...metadata, documentTitle: "Reply" },
      }),
      (error) => error.code === "DOCUMENT_EXISTS",
    );
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});

test("project clean export publishes a versioned DOCX without overwriting", async () => {
  const directory = await mkdtemp(join(tmpdir(), "agent-docx-export-"));
  const manifestPath = join(directory, "agent-docx.json");
  const sourcePath = join(directory, "motion.md");
  const outputPath = join(directory, "motion.docx");
  try {
    await writeFile(sourcePath, "# Motion\n\nThe requested relief follows.\n");
    const project = await createProject(manifestPath, {
      documentId: "motion",
      source: "motion.md",
      profile: "us-district-conventional",
      metadata,
    });
    const checkpoint = await project.checkpoint("motion", {
      baseRevision: null,
      author: { name: "Drafter" },
      message: "Initial draft",
    });
    const exported = await project.exportDocx("motion", {
      revision: checkpoint.revision.id,
      mode: "clean",
      output: outputPath,
    });
    assert.equal(exported.artifact.revision, checkpoint.revision.id);
    assert.equal(exported.artifact.path, outputPath);
    assert.equal((await readFile(outputPath)).byteLength, exported.bytes.byteLength);
    const inspected = await inspectDocxMaterial(exported.bytes);
    assert.equal(inspected.semantic?.revision, checkpoint.revision.id);
    assert.equal(inspected.semantic?.baseRevision, null);
    assert.ok(
      inspected.semantic?.dependencies.some(
        (dependency) => dependency.key === "profile",
      ),
    );
    await assert.rejects(
      project.exportDocx("motion", {
        revision: checkpoint.revision.id,
        mode: "clean",
        output: outputPath,
      }),
      (error) => error.code === "OUTPUT_EXISTS",
    );
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});

test("project export publishes exhibit attachments beside the DOCX", async () => {
  const directory = await mkdtemp(join(tmpdir(), "agent-docx-attachments-"));
  const manifestPath = join(directory, "agent-docx.json");
  const sourcePath = join(directory, "motion.md");
  const assetsPath = join(directory, "assets");
  const outputPath = join(directory, "motion.docx");
  const record = Buffer.from("%PDF-1.7\nrecord\n");
  try {
    await mkdir(assetsPath);
    await writeFile(join(assetsPath, "record.pdf"), record);
    await writeFile(
      sourcePath,
      ':::exhibit{id="exhibit-a" label="Exhibit A" source="record.pdf"}\n\nAttached record.\n\n:::\n',
    );
    const project = await createProject(manifestPath, {
      documentId: "motion",
      source: "motion.md",
      assetsDir: "assets",
      profile: "us-district-conventional",
      metadata,
    });
    const checkpoint = await project.checkpoint("motion", {
      baseRevision: null,
      author: { name: "Drafter" },
      message: "Initial draft",
    });
    const exported = await project.exportDocx("motion", {
      revision: checkpoint.revision.id,
      mode: "clean",
      output: outputPath,
    });
    const bundlePath = join(directory, "motion.attachments");
    assert.equal(exported.artifact.attachments?.path, bundlePath);
    assert.equal(
      exported.artifact.attachments?.manifest.entries[0].name,
      "record.pdf",
    );
    assert.deepEqual(
      await readFile(join(bundlePath, "files", "record.pdf")),
      record,
    );
    assert.equal(
      JSON.parse(await readFile(join(bundlePath, "manifest.json"), "utf8"))
        .entries[0].payloadPath,
      "files/record.pdf",
    );
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});

test("project redline export carries native insertions and deletions", async () => {
  const directory = await mkdtemp(join(tmpdir(), "agent-docx-redline-"));
  const manifestPath = join(directory, "agent-docx.json");
  const sourcePath = join(directory, "motion.md");
  const outputPath = join(directory, "motion-redline.docx");
  try {
    await writeFile(sourcePath, "# Motion\n\nOld statement.\n");
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
    const baseDocument = await project.getDocument("motion", base.revision.id);
    const body = baseDocument.document.blocks[1];
    assert.equal(body?.kind, "paragraph");
    const reviewed = await project.addReview("motion", {
      revision: base.revision.id,
      blockId: body.id,
      author: { name: "Reviewer" },
      message: "Confirm the record citation.",
    });
    const reviewAnnotation = (
      await project.getDocument("motion", reviewed.revision.id)
    ).annotations[0];
    assert.ok(reviewAnnotation);
    await writeFile(
      sourcePath,
      (await readFile(sourcePath, "utf8")).replace("Old statement.", "New statement."),
    );
    const head = await project.checkpoint("motion", {
      baseRevision: reviewed.revision.id,
      author: { name: "Drafter" },
      message: "Revise statement",
    });
    const exported = await project.exportDocx("motion", {
      revision: head.revision.id,
      mode: "redline",
      baseRevision: base.revision.id,
      output: outputPath,
    });
    assert.equal(exported.artifact.mode, "redline");
    const parts = await readDocxParts(exported.bytes);
    const documentXml = new TextDecoder().decode(parts.get("word/document.xml"));
    const settingsXml = new TextDecoder().decode(parts.get("word/settings.xml"));
    const commentsXml = new TextDecoder().decode(parts.get("word/comments.xml"));
    assert.match(documentXml, /<w:ins /);
    assert.match(documentXml, /<w:del /);
    assert.match(settingsXml, /<w:trackRevisions\/>/);
    assert.match(documentXml, /<w:commentRangeStart /);
    assert.match(documentXml, /<w:commentReference /);
    assert.match(commentsXml, /Confirm the record citation\./);
    const inspected = await inspectDocxMaterial(exported.bytes);
    assert.equal(inspected.semantic?.mode, "redline");
    assert.equal(inspected.semantic?.baseRevision, base.revision.id);
    assert.equal(inspected.semantic?.revision, head.revision.id);
    assert.equal(inspected.semantic?.revisionMap.length, 1);
    assert.equal(inspected.semantic?.commentMap.length, 1);
    assert.equal(inspected.semantic?.commentMap[0]?.blockWide, true);
    assert.deepEqual(inspected.result.recognized.annotations, [reviewAnnotation]);
    assert.deepEqual(exported.artifact.rendererProvenance.verification, {
      revisionCount: 2,
      commentCount: 1,
      fieldCount: 0,
    });
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});

test("strict project import reconstructs a generated text redline", async () => {
  const sourceDirectory = await mkdtemp(join(tmpdir(), "agent-docx-redline-source-"));
  const targetDirectory = await mkdtemp(join(tmpdir(), "agent-docx-redline-target-"));
  const sourceManifest = join(sourceDirectory, "agent-docx.json");

  const targetManifest = join(targetDirectory, "agent-docx.json");
  const sourcePath = join(sourceDirectory, "motion.md");
  const targetPath = join(targetDirectory, "motion.md");
  const projectId = "00000000-0000-4000-8000-000000000001";
  try {
    await writeFile(sourcePath, "# Motion\n\nOld statement.\n");
    const sourceProject = await createProject(
      sourceManifest,
      {
        documentId: "motion",
        source: "motion.md",
        profile: "us-district-conventional",
        metadata,
      },
      { randomUUID: () => projectId },
    );
    const base = await sourceProject.checkpoint("motion", {
      baseRevision: null,
      author: { name: "Drafter" },
      message: "Initial draft",
    });
    const baseDocument = await sourceProject.getDocument("motion", base.revision.id);
    const baseBody = baseDocument.document.blocks[1];
    assert.equal(baseBody?.kind, "paragraph");
    const reviewed = await sourceProject.addReview("motion", {
      revision: base.revision.id,
      blockId: baseBody.id,
      author: { name: "Reviewer" },
      message: "Confirm the record citation.",
    });
    const reviewAnnotation = (
      await sourceProject.getDocument("motion", reviewed.revision.id)
    ).annotations[0];
    assert.ok(reviewAnnotation);
    await writeFile(
      sourcePath,
      (await readFile(sourcePath, "utf8")).replace("Old statement.", "New statement."),
    );
    const head = await sourceProject.checkpoint("motion", {
      baseRevision: reviewed.revision.id,
      author: { name: "Drafter" },
      message: "Revise statement",
    });
    const exported = await sourceProject.exportDocx("motion", {
      revision: head.revision.id,
      mode: "redline",
      baseRevision: base.revision.id,
      output: join(sourceDirectory, "motion-redline.docx"),
    });
    const targetProject = await createProject(
      targetManifest,
      {
        documentId: "motion",
        source: "motion.md",
        createSource: true,
        profile: "us-district-conventional",
        metadata,
      },
      { randomUUID: () => projectId },
    );
    const imported = await targetProject.importDocx({
      input: exported.bytes,
      inspectOnly: false,
      documentId: "motion",
      output: targetPath,
      author: { name: "Importer" },
      message: "Import redline",
    });
    assert.equal(imported.mode, "tracked");
    assert.equal(imported.revisions.length, 2);
    assert.match(await readFile(targetPath, "utf8"), /New statement\./);
    assert.match(
      (await targetProject.getDocument("motion", imported.baseRevision)).source,
      /Old statement\./,
    );
    const importedHead = await targetProject.getDocument(
      "motion",
      imported.headRevision,
    );
    assert.deepEqual(importedHead.document.annotations, [reviewAnnotation]);
  } finally {
    await rm(sourceDirectory, { recursive: true, force: true });
    await rm(targetDirectory, { recursive: true, force: true });
  }
});
test("review-only redline preserves comments without synthetic revisions", async () => {
  const directory = await mkdtemp(join(tmpdir(), "agent-docx-review-redline-"));
  const manifestPath = join(directory, "agent-docx.json");
  const sourcePath = join(directory, "motion.md");
  try {
    await writeFile(sourcePath, "# Motion\n\nReview this statement.\n");
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
    const document = await project.getDocument("motion", base.revision.id);
    const body = document.document.blocks[1];
    assert.equal(body?.kind, "paragraph");
    const reviewed = await project.addReview("motion", {
      revision: base.revision.id,
      blockId: body.id,
      author: { name: "Reviewer" },
      message: "Confirm the record citation.",
    });
    const annotation = (
      await project.getDocument("motion", reviewed.revision.id)
    ).annotations[0];
    assert.ok(annotation);
    const exported = await project.exportDocx("motion", {
      revision: reviewed.revision.id,
      mode: "redline",
      baseRevision: base.revision.id,
      output: join(directory, "motion-redline.docx"),
    });
    const inspected = await inspectDocxMaterial(exported.bytes);
    assert.equal(inspected.tracked?.baseSource, inspected.tracked?.headSource);
    assert.deepEqual(inspected.result.recognized.annotations, [annotation]);
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});

test("strict project import establishes a root revision from generated DOCX", async () => {
  const directory = await mkdtemp(join(tmpdir(), "agent-docx-import-"));
  const manifestPath = join(directory, "agent-docx.json");
  const sourcePath = join(directory, "motion.md");
  try {
    const project = await createProject(manifestPath, {
      documentId: "motion",
      source: "motion.md",
      createSource: true,
      profile: "us-district-conventional",
      metadata,
    });
    const compiled = await compileMarkdown("Imported paragraph.\n", {
      projectId: (await project.getState()).manifest.projectId,
      documentId: "motion",
      profile: "us-district-conventional",
      metadata,
    });
    const imported = await project.importDocx({
      input: compiled.bytes,
      inspectOnly: false,
      documentId: "motion",
      output: sourcePath,
      author: { name: "Importer" },
      message: "Import DOCX",
    });
    assert.equal(imported.mode, "clean");
    assert.equal(imported.revisions.length, 1);
    assert.equal((await project.getDocument("motion", "HEAD")).document.blocks.length, 1);
    assert.match(await readFile(sourcePath, "utf8"), /agent-docx:block/);
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});

test("strict project import materializes semantic embedded assets", async () => {
  const directory = await mkdtemp(join(tmpdir(), "agent-docx-import-assets-"));
  const manifestPath = join(directory, "agent-docx.json");
  const sourcePath = join(directory, "motion.md");
  const assetsPath = join(directory, "assets");
  const seal = Uint8Array.from(
    Buffer.from(
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4z8DwHwAFgAI/ScL77QAAAABJRU5ErkJggg==",
      "base64",
    ),
  );
  try {
    await mkdir(assetsPath);
    const project = await createProject(manifestPath, {
      documentId: "motion",
      source: "motion.md",
      createSource: true,
      assetsDir: "assets",
      profile: "us-district-conventional",
      metadata,
    });
    const compiled = await compileMarkdown(
      '::image{source="seal.png" alt="Court seal" widthTwips="1440" heightTwips="1440"}\n',
      {
        projectId: (await project.getState()).manifest.projectId,
        documentId: "motion",
        profile: "us-district-conventional",
        metadata,
        assets: { "seal.png": { bytes: seal, mediaType: "image/png" } },
      },
    );
    await project.importDocx({
      input: compiled.bytes,
      inspectOnly: false,
      documentId: "motion",
      output: sourcePath,
      author: { name: "Importer" },
      message: "Import DOCX assets",
    });
    assert.deepEqual(await readFile(join(assetsPath, "seal.png")), Buffer.from(seal));
    assert.equal((await project.getDocument("motion", "HEAD")).document.assets["seal.png"].mediaType, "image/png");
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});

test("strict project import consumes an authorized exhibit attachment bundle", async () => {
  const sourceDirectory = await mkdtemp(join(tmpdir(), "agent-docx-exhibit-source-"));
  const targetDirectory = await mkdtemp(join(tmpdir(), "agent-docx-exhibit-target-"));
  const sourceManifest = join(sourceDirectory, "agent-docx.json");
  const sourcePath = join(sourceDirectory, "motion.md");
  const sourceAssets = join(sourceDirectory, "assets");
  const outputPath = join(sourceDirectory, "motion.docx");
  const targetManifest = join(targetDirectory, "agent-docx.json");
  const targetPath = join(targetDirectory, "motion.md");
  const targetAssets = join(targetDirectory, "assets");
  const record = Buffer.from("%PDF-1.7\nrecord\n");
  try {
    await mkdir(sourceAssets);
    await mkdir(targetAssets);
    await writeFile(join(sourceAssets, "record.pdf"), record);
    await writeFile(
      sourcePath,
      ':::exhibit{id="exhibit-a" label="Exhibit A" source="record.pdf"}\n\nAttached record.\n\n:::\n',
    );
    const sourceProject = await createProject(sourceManifest, {
      documentId: "motion",
      source: "motion.md",
      assetsDir: "assets",
      profile: "us-district-conventional",
      metadata,
    });
    const sourceState = await sourceProject.getState();
    const checkpoint = await sourceProject.checkpoint("motion", {
      baseRevision: null,
      author: { name: "Drafter" },
      message: "Initial draft",
    });
    const exported = await sourceProject.exportDocx("motion", {
      revision: checkpoint.revision.id,
      mode: "clean",
      output: outputPath,
    });
    const targetProject = await createProject(
      targetManifest,
      {
        documentId: "motion",
        source: "motion.md",
        createSource: true,
        assetsDir: "assets",
        profile: "us-district-conventional",
        metadata,
      },
      { randomUUID: () => sourceState.manifest.projectId },
    );
    const imported = await targetProject.importDocx({
      input: exported.bytes,
      attachments: { directory: join(sourceDirectory, "motion.attachments") },
      inspectOnly: false,
      documentId: "motion",
      output: targetPath,
      author: { name: "Importer" },
      message: "Import exhibit",
    });
    assert.equal(imported.mode, "clean");
    assert.deepEqual(await readFile(join(targetAssets, "record.pdf")), record);
    assert.match(await readFile(targetPath, "utf8"), /Exhibit A/);
  } finally {
    await rm(sourceDirectory, { recursive: true, force: true });
    await rm(targetDirectory, { recursive: true, force: true });
  }
});

test("strict clean import reconstructs native comments", async () => {
  const sourceDirectory = await mkdtemp(join(tmpdir(), "agent-docx-clean-comments-source-"));
  const targetDirectory = await mkdtemp(join(tmpdir(), "agent-docx-clean-comments-target-"));
  const sourceManifest = join(sourceDirectory, "agent-docx.json");
  const sourcePath = join(sourceDirectory, "motion.md");
  const outputPath = join(sourceDirectory, "motion.docx");
  const targetManifest = join(targetDirectory, "agent-docx.json");
  const targetPath = join(targetDirectory, "motion.md");
  try {
    await writeFile(sourcePath, "# Motion\n\nBody.\n");
    const sourceProject = await createProject(sourceManifest, {
      documentId: "motion",
      source: "motion.md",
      profile: "us-district-conventional",
      metadata,
    });
    const sourceState = await sourceProject.getState();
    const checkpoint = await sourceProject.checkpoint("motion", {
      baseRevision: null,
      author: { name: "Drafter" },
      message: "Initial draft",
    });
    const exported = await sourceProject.exportDocx("motion", {
      revision: checkpoint.revision.id,
      mode: "clean",
      output: outputPath,
    });
    const parts = await readDocxParts(exported.bytes);
    const documentXml = new TextDecoder().decode(parts.get("word/document.xml"));
    const markedDocumentXml = documentXml.replace(
      /(<w:p\b[^>]*>)([\s\S]*?)(<\/w:p>)/,
      (_, open, body, close) => {
        const firstRun = body.indexOf("<w:r");
        const lastRunEnd = body.lastIndexOf("</w:r>") + "</w:r>".length;
        assert.ok(firstRun >= 0 && lastRunEnd > firstRun);
        return `${open}<w:commentRangeStart w:id="0"/>${body.slice(
          0,
          firstRun,
        )}${body.slice(firstRun, lastRunEnd)}<w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r>${body.slice(
          lastRunEnd,
        )}${close}`;
      },
    );
    assert.notEqual(markedDocumentXml, documentXml);
    const modified = new Map(parts);
    modified.set("word/document.xml", new TextEncoder().encode(markedDocumentXml));
    modified.set(
      "word/comments.xml",
      new TextEncoder().encode(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="0" w:author="Reviewer" w:date="2026-08-02T00:00:00Z"><w:p><w:r><w:t>Check title.</w:t></w:r></w:p></w:comment></w:comments>',
      ),
    );
    const targetProject = await createProject(
      targetManifest,
      {
        documentId: "motion",
        source: "motion.md",
        createSource: true,
        profile: "us-district-conventional",
        metadata,
      },
      { randomUUID: () => sourceState.manifest.projectId },
    );
    const imported = await targetProject.importDocx({
      input: repackDocxParts(modified),
      inspectOnly: false,
      documentId: "motion",
      output: targetPath,
      author: { name: "Importer" },
      message: "Import comment",
    });
    assert.equal(imported.mode, "clean");
    assert.equal(imported.recognized.annotations.length, 1);
    assert.equal(imported.recognized.annotations[0].message, "Check title.");
    assert.equal((await targetProject.getDocument("motion", "HEAD")).annotations.length, 1);
  } finally {
    await rm(sourceDirectory, { recursive: true, force: true });
    await rm(targetDirectory, { recursive: true, force: true });
  }
});

test("project evaluates and applies canonical, Unicode-safe source patches", async () => {
  const directory = await mkdtemp(join(tmpdir(), "agent-docx-patch-"));
  const manifestPath = join(directory, "agent-docx.json");
  const sourcePath = join(directory, "motion.md");
  try {
    await writeFile(sourcePath, "# Motion\n\nA 😀 body.\n");
    const project = await createProject(manifestPath, {
      documentId: "motion",
      source: "motion.md",
      profile: "us-district-conventional",
      metadata,
    });
    const initial = await project.checkpoint("motion", {
      baseRevision: null,
      author: { name: "Drafter" },
      message: "Initial draft",
    });
    const guidance = await project.getDraftGuidance("motion", initial.revision.id);
    assert.equal(guidance.baseRevision, initial.revision.id);
    assert.ok(guidance.items.length > 0);
    const marked = await readFile(sourcePath, "utf8");
    const bodyStart = marked.indexOf("body");
    const validPatch = {
      schemaVersion: 1,
      documentId: "motion",
      baseRevision: initial.revision.id,
      edits: [{
        start: bodyStart,
        end: bodyStart + "body".length,
        expectedText: "body",
        replacement: "submission",
      }],
    };
    const evaluation = await project.evaluatePatch(validPatch);
    assert.equal(evaluation.candidate.status, "ok");
    assert.equal(evaluation.canApply, true);
    const applied = await project.applyPatch(validPatch, {
      patchHash: evaluation.patchHash,
      gate: "report",
      author: { name: "Drafter" },
      message: "Expand body",
    });
    assert.notEqual(applied.revision.id, initial.revision.id);
    assert.match(await readFile(sourcePath, "utf8"), /submission/);

    const current = await readFile(sourcePath, "utf8");
    const emoji = current.indexOf("😀");
    const unsafePatch = {
      schemaVersion: 1,
      documentId: "motion",
      baseRevision: applied.revision.id,
      edits: [{
        start: emoji + 1,
        end: emoji + 1,
        expectedText: "",
        replacement: "x",
      }],
    };
    const unsafe = await project.evaluatePatch(unsafePatch);
    assert.equal(unsafe.candidate.status, "invalid");
    assert.equal(unsafe.candidate.error.code, "PATCH_INVALID");
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});
