import assert from "node:assert/strict";
import { createHash, randomUUID } from "node:crypto";
import { cp, mkdir, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { createProject } from "../dist/index.js";
import { decodeDocxXml, readDocxParts } from "../dist/docx/package.js";

const root = dirname(fileURLToPath(import.meta.url));
const fixture = join(root, "..", "test", "fixtures", "agent-project");
const hash = (bytes) =>
  `sha256:${createHash("sha256").update(bytes).digest("hex")}`;
const readJson = async (path) => JSON.parse(await readFile(path, "utf8"));
const allXml = async (bytes) => {
  const parts = await readDocxParts(bytes);
  return [...parts.entries()]
    .filter(([path]) => path.endsWith(".xml") || path.endsWith(".rels"))
    .map(([path, value]) => [path, decodeDocxXml(value)])
    .sort(([left], [right]) => (left < right ? -1 : left > right ? 1 : 0));
};
const assertArtifact = async (compiled, label) => {
  assert.ok(compiled.bytes.byteLength > 0, `${label} DOCX is empty`);
  assert.equal(compiled.artifact.byteLength, compiled.bytes.byteLength);
  assert.equal(compiled.artifact.sha256, hash(compiled.bytes));
  assert.ok(compiled.artifact.path);
  assert.ok(compiled.artifact.storePath);
  assert.equal(
    hash(await readFile(compiled.artifact.path)),
    compiled.artifact.sha256,
  );
  assert.notEqual(compiled.artifact.storePath, compiled.artifact.path);
  if (compiled.attachments) {
    assert.ok(compiled.artifact.attachments?.path);
    assert.ok(compiled.artifact.attachments?.storePath);
    assert.equal(
      compiled.artifact.attachments.manifestSha256,
      compiled.attachments.manifestSha256,
    );
    for (const entry of compiled.attachments.manifest.entries) {
      const attachment = compiled.attachments.files[entry.name];
      assert.ok(attachment, `${label} attachment missing: ${entry.name}`);
      assert.equal(attachment.bytes.byteLength, entry.byteLength);
      assert.equal(hash(attachment.bytes), entry.sha256);
      assert.equal(
        hash(
          await readFile(
            join(compiled.artifact.attachments.path, entry.payloadPath),
          ),
        ),
        entry.sha256,
      );
    }
  }
};

const main = async () => {
  const directory = await mkdtemp(join(tmpdir(), "agent-docx-smoke-"));
  const manifestPath = join(directory, "agent-docx.json");
  const sourcePath = join(directory, "motion.md");
  const assetsPath = join(directory, "assets");
  const cleanPath = join(directory, "motion-clean.docx");
  const redlinePath = join(directory, "motion-redline.docx");
  const targetDirectory = await mkdtemp(
    join(tmpdir(), "agent-docx-smoke-import-"),
  );
  try {
    await mkdir(assetsPath);
    await cp(join(fixture, "metadata.json"), join(directory, "metadata.json"));
    await cp(join(fixture, "chrome.json"), join(directory, "chrome.json"));
    const metadata = await readJson(join(directory, "metadata.json"));
    const chrome = await readJson(join(directory, "chrome.json"));
    const baseSource = await readFile(join(fixture, "motion.md"), "utf8");
    const record = Buffer.from("%PDF-1.7\nagent-docx smoke exhibit\n");
    await writeFile(join(assetsPath, "record.pdf"), record);
    await writeFile(
      sourcePath,
      `${baseSource}\n:::exhibit{id="exhibit-a" label="Exhibit A" source="record.pdf"}\n\nAttached record.\n\n:::\n`,
    );

    const project = await createProject(manifestPath, {
      documentId: "motion",
      source: "motion.md",
      assetsDir: "assets",
      profile: "cand-civil",
      filingKind: "motion-document",
      rulePack: "cand-civil@2026-05-01",
      metadata,
      chrome,
    });
    const initialized = await project.getState();
    const initial = await project.checkpoint("motion", {
      baseRevision: null,
      author: { name: "Drafter" },
      message: "Initial smoke draft",
    });
    const baseRevision = initial.revision.id;
    const base = await project.getDocument("motion", baseRevision);
    const insertedBlockId = `b_${randomUUID()}`;
    const insertedText =
      "The requested relief should be granted without a narrower remedy.";
    const replacement =
      "the record establishes the required elements and no narrower remedy will cure the injury.";
    const replacementTarget = "the record establishes the required elements.";
    const insert = `\n\n<!-- agent-docx:block id="${insertedBlockId}" -->\n${insertedText}\n`;
    const replacementStart = base.source.indexOf(replacementTarget);
    assert.ok(
      replacementStart >= 0,
      "smoke fixture replacement target missing",
    );
    const patch = {
      schemaVersion: 1,
      documentId: "motion",
      baseRevision,
      edits: [
        {
          start: replacementStart,
          end: replacementStart + replacementTarget.length,
          expectedText: replacementTarget,
          replacement,
        },
        {
          start: base.source.length,
          end: base.source.length,
          expectedText: "",
          replacement: insert,
        },
      ],
    };
    const evaluation = await project.evaluatePatch(patch);
    assert.equal(evaluation.candidate.status, "ok");
    assert.equal(evaluation.canApply, true);
    assert.equal(evaluation.passesConstraints, false);
    const unknownFindings = evaluation.candidate.validation.findings.filter(
      (finding) => finding.status === "unknown",
    );
    assert.equal(unknownFindings.length, 1);
    assert.equal(unknownFindings[0].severity, "error");
    assert.match(
      `${unknownFindings[0].checkId} ${unknownFindings[0].message}`,
      /pitch|typeface/i,
    );
    assert.equal(evaluation.candidate.validation.scope.certification, false);

    const applied = await project.applyPatch(patch, {
      patchHash: evaluation.patchHash,
      gate: "not-worse",
      author: { name: "Drafter" },
      message: "Apply mixed smoke patch",
    });
    assert.equal(applied.revision.parents.length, 1);
    const patched = await project.getDocument("motion", applied.revision.id);
    const insertedBlock = patched.document.blocks.find(
      (block) => block.id === insertedBlockId,
    );
    assert.ok(insertedBlock, "inserted fixed-marker block missing");
    const reviewBlock = patched.document.blocks.find(
      (block) => block.id !== insertedBlockId && block.kind === "heading",
    );
    assert.ok(reviewBlock, "unchanged review block missing");
    const reviewText = reviewBlock.runs.map((run) => run.text).join("");
    const reviewEnd = Math.min(4, reviewText.length);
    const reviewed = await project.addReview("motion", {
      revision: applied.revision.id,
      blockId: reviewBlock.id,
      range: { start: 0, end: reviewEnd },
      author: { name: "Reviewer" },
      message: "Review unchanged heading",
    });
    assert.equal(reviewed.revision.parents.length, 1);
    const headRevision = reviewed.revision.id;
    const head = await project.getDocument("motion", headRevision);
    assert.equal(head.annotations.length, 1);
    assert.deepEqual(head.annotations[0].range, { start: 0, end: reviewEnd });
    const changes = await project.diff("motion", baseRevision, headRevision);
    assert.ok(changes.changes.length > 0);
    assert.ok(changes.annotations.length > 0);

    const clean = await project.exportDocx("motion", {
      revision: headRevision,
      mode: "clean",
      output: cleanPath,
    });
    const redline = await project.exportDocx("motion", {
      revision: headRevision,
      mode: "redline",
      baseRevision,
      output: redlinePath,
    });
    await assertArtifact(clean, "clean");
    await assertArtifact(redline, "redline");
    assert.ok(
      (redline.artifact.rendererProvenance.verification?.revisionCount ?? 0) >
        0,
    );
    assert.ok(
      (redline.artifact.rendererProvenance.verification?.commentCount ?? 0) > 0,
    );
    const cleanXml = await allXml(clean.bytes);
    const redlineXml = await allXml(redline.bytes);
    const cleanText = cleanXml.map(([, xml]) => xml).join("\n");
    const redlineText = redlineXml.map(([, xml]) => xml).join("\n");
    assert.match(cleanText, /Example Holdings/);
    assert.match(cleanText, /Page/);
    assert.match(cleanText, /NUMPAGES|PAGE/);
    assert.match(redlineText, /w:ins/);
    assert.match(redlineText, /w:del/);
    assert.match(redlineText, /w:commentRangeStart/);
    assert.match(cleanText, /exhibit-a|Exhibit A/);

    const targetManifest = join(targetDirectory, "agent-docx.json");
    const targetSource = join(targetDirectory, "motion.md");
    const targetAssets = join(targetDirectory, "assets");
    await mkdir(targetAssets);
    const target = await createProject(
      targetManifest,
      {
        documentId: "motion",
        source: "motion.md",
        createSource: true,
        assetsDir: "assets",
        profile: "cand-civil",
        filingKind: "motion-document",
        rulePack: "cand-civil@2026-05-01",
        metadata,
        chrome,
      },
      { randomUUID: () => initialized.manifest.projectId },
    );
    const imported = await target.importDocx({
      input: clean.bytes,
      inspectOnly: false,
      attachments: { directory: clean.artifact.attachments.path },
      documentId: "motion",
      output: targetSource,
      author: { name: "Importer" },
      message: "Re-import generated clean DOCX",
    });
    assert.equal(imported.mode, "clean");
    assert.equal(imported.fidelity.overall, "normalized");
    assert.equal(
      imported.fidelity.items.some((item) => item.status === "unsupported"),
      false,
    );
    const importedHead = await target.getDocument(
      "motion",
      imported.headRevision,
    );
    assert.equal(
      importedHead.document.blocks.length,
      head.document.blocks.length,
    );
    assert.equal(
      importedHead.document.blocks[0].id,
      head.document.blocks[0].id,
    );
    assert.deepEqual(importedHead.document.metadata, head.document.metadata);
    assert.deepEqual(importedHead.document.chrome, head.document.chrome);
    assert.deepEqual(await readFile(join(targetAssets, "record.pdf")), record);
    const importedMeasurement = await target.measure(
      "motion",
      imported.headRevision,
    );
    assert.equal(
      importedMeasurement.deterministic.totalVisualLines,
      clean.measurement.deterministic.totalVisualLines,
    );
    assert.equal(importedMeasurement.pageCount, clean.measurement.pageCount);

    const serializable = JSON.stringify({
      initialized,
      evaluation: {
        ...evaluation,
        candidate:
          evaluation.candidate.status === "ok"
            ? { ...evaluation.candidate, measurement: undefined }
            : evaluation.candidate,
      },
      clean: {
        ...clean,
        bytes: undefined,
        attachments: clean.attachments
          ? {
              manifest: clean.attachments.manifest,
              manifestSha256: clean.attachments.manifestSha256,
            }
          : null,
      },
      redline: {
        ...redline,
        bytes: undefined,
        attachments: redline.attachments
          ? {
              manifest: redline.attachments.manifest,
              manifestSha256: redline.attachments.manifestSha256,
            }
          : null,
      },
      imported,
    });
    assert.ok(!serializable.includes("Uint8Array"));
    assert.ok(!serializable.includes('"bytes":{"0"'));
    console.log(
      JSON.stringify({
        project: manifestPath,
        baseRevision,
        appliedRevision: applied.revision.id,
        reviewRevision: headRevision,
        clean: clean.artifact,
        redline: redline.artifact,
        importedRevision: imported.headRevision,
      }),
    );
  } finally {
    await rm(directory, { recursive: true, force: true });
    await rm(targetDirectory, { recursive: true, force: true });
  }
};

await main();
