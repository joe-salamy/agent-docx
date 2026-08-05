import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";
import { createProject } from "../dist/index.js";
import { inspectDocxMaterial } from "../dist/docx/import.js";
import { resolveLibreOffice } from "../dist/renderers/office.js";
import { metadata } from "./helpers.js";

const sha256 = (bytes) =>
  `sha256:${createHash("sha256").update(bytes).digest("hex")}`;

let libreOfficeAvailable = false;
try {
  await resolveLibreOffice();
  libreOfficeAvailable = true;
} catch {}

test(
  "project PDF export publishes a verified PDF and DOCX artifact",
  { skip: !libreOfficeAvailable },
  async () => {
    const directory = await mkdtemp(join(tmpdir(), "agent-docx-pdf-export-"));
    const manifestPath = join(directory, "agent-docx.json");
    const sourcePath = join(directory, "motion.md");
    const outputPath = join(directory, "motion.pdf");
    try {
      await writeFile(
        sourcePath,
        "# Motion\n\nThe requested relief follows.\n",
      );
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
        mode: "pdf",
        output: outputPath,
      });
      const outputBytes = await readFile(outputPath);
      assert.ok(outputBytes.byteLength > 0);
      assert.equal(outputBytes.subarray(0, 4).toString(), "%PDF");
      assert.ok(exported.artifact.pdf);
      assert.equal(exported.artifact.pdf.path, outputPath);
      assert.equal(exported.artifact.pdf.sha256, sha256(outputBytes));
      assert.equal(
        exported.artifact.pdf.delta,
        exported.artifact.pdf.pageCount -
          exported.artifact.pdf.deterministicPageCount,
      );
      assert.equal(Number.isFinite(exported.artifact.pdf.delta), true);
      const storedPdf = await readFile(exported.artifact.pdf.storePath);
      assert.equal(sha256(storedPdf), exported.artifact.pdf.sha256);
      const storedDocx = await readFile(exported.artifact.storePath);
      const inspected = await inspectDocxMaterial(storedDocx);
      assert.notEqual(inspected.result.fidelity.overall, "unsupported");
      await assert.rejects(
        () =>
          project.exportDocx("motion", {
            revision: checkpoint.revision.id,
            mode: "pdf",
            output: outputPath,
          }),
        (error) => error?.code === "OUTPUT_EXISTS",
      );
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  },
);

test("PDF export reports a missing explicit LibreOffice executable", async () => {
  const directory = await mkdtemp(join(tmpdir(), "agent-docx-pdf-missing-lo-"));
  const manifestPath = join(directory, "agent-docx.json");
  const sourcePath = join(directory, "motion.md");
  const outputPath = join(directory, "motion.pdf");
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
    await assert.rejects(
      () =>
        project.exportDocx("motion", {
          revision: checkpoint.revision.id,
          mode: "pdf",
          output: outputPath,
          options: {
            libreoffice: { executablePath: join(directory, "missing-soffice") },
          },
        }),
      (error) => error?.code === "LIBREOFFICE_NOT_FOUND",
    );
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});
