import test from "node:test";
import assert from "node:assert/strict";
import { chmod, mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { createHash } from "node:crypto";
import { PDFDocument } from "pdf-lib";
import {
  renderLibreOffice,
  renderWord,
  resolveLibreOffice,
} from "../dist/renderers/office.js";
import { AgentDocxError } from "../dist/index.js";
import { measureMarkdown } from "../dist/index.js";

test("LibreOffice resolver rejects an explicit missing executable", async () => {
  await assert.rejects(
    () => resolveLibreOffice("/definitely/missing/soffice"),
    (error) =>
      error instanceof AgentDocxError &&
      error.code === "LIBREOFFICE_NOT_FOUND",
  );
});

test("explicit renderer paths must be absolute", async () => {
  await assert.rejects(
    () => resolveLibreOffice("relative-soffice"),
    (error) =>
      error instanceof AgentDocxError && error.code === "INVALID_ARGUMENT",
  );
  await assert.rejects(
    () =>
      renderWord(new Uint8Array([1]), [], {
        powerShellPath: "relative-powershell",
      }),
    (error) =>
      error instanceof AgentDocxError && error.code === "INVALID_ARGUMENT",
  );
});

test("LibreOffice terminates runaway transport output", async () => {
  const root = await mkdtemp(join(tmpdir(), "agent-docx-overflow-lo-"));
  const executable = join(root, "soffice");
  const script = `#!/usr/bin/env node
const chunk = "x".repeat(65536);
function write() {
  while (process.stdout.write(chunk)) {}
  process.stdout.once("drain", write);
}
write();`;
  try {
    await writeFile(executable, script);
    await chmod(executable, 0o755);
    const started = performance.now();
    await assert.rejects(
      () =>
        renderLibreOffice(new Uint8Array([1]), [], {
          executablePath: executable,
        }),
      (error) =>
        error instanceof AgentDocxError &&
        error.code === "LIBREOFFICE_RENDER_FAILED" &&
        /transport limit/.test(error.message),
    );
    assert.ok(
      performance.now() - started < 8000,
      "transport overflow did not terminate the child promptly",
    );
  } finally {
    await rm(root, { recursive: true, force: true });
  }
});

test("LibreOffice adapter uses isolated exact conversion arguments and PDF page tree", async () => {
  const root = await mkdtemp(join(tmpdir(), "agent-docx-fake-lo-"));
  const executable = join(root, "soffice");
  const pdf = await PDFDocument.create();
  pdf.addPage();
  pdf.addPage();
  process.env.AGENT_DOCX_FAKE_PDF = Buffer.from(await pdf.save()).toString(
    "base64",
  );
  const script = `#!/usr/bin/env node\nimport {writeFileSync} from 'node:fs';import {join} from 'node:path';const a=process.argv.slice(2);if(a[0]==='--version'){console.log('LibreOffice 99.0 fake');process.exit(0)}const expected=['--headless','--nologo','--nodefault','--norestore','--infilter=Office Open XML Text','--convert-to','pdf:writer_pdf_Export'];for(const x of expected)if(!a.includes(x)){console.error('missing '+x);process.exit(9)}const out=a[a.indexOf('--outdir')+1];writeFileSync(join(out,'render.pdf'),Buffer.from(process.env.AGENT_DOCX_FAKE_PDF,'base64'));`;
  try {
    await writeFile(executable, script);
    await chmod(executable, 0o755);
    const rendered = await renderLibreOffice(
      new Uint8Array([1, 2, 3]),
      ["Times New Roman"],
      { executablePath: executable },
      10000,
    );
    assert.equal(rendered.pageCount, 2);
    assert.match(rendered.versionRaw, /99.0 fake/);
    assert.equal(rendered.calibratedFontEnvironment, false);
    assert.equal(rendered.requestedFontFamilies[0], "Times New Roman");
  } finally {
    delete process.env.AGENT_DOCX_FAKE_PDF;
    await rm(root, { recursive: true, force: true });
  }
});

test("LibreOffice renders the exact returned generated DOCX", async () => {
  const root = await mkdtemp(join(tmpdir(), "agent-docx-fake-lo-sha-"));
  const executable = join(root, "soffice");
  const pdf = await PDFDocument.create();
  pdf.addPage();
  process.env.AGENT_DOCX_FAKE_PDF = Buffer.from(await pdf.save()).toString(
    "base64",
  );
  const script = `#!/usr/bin/env node\nimport {writeFileSync} from 'node:fs';import {join} from 'node:path';const a=process.argv.slice(2);if(a[0]==='--version'){console.log('LibreOffice 99.0 fake');process.exit(0)}const out=a[a.indexOf('--outdir')+1];writeFileSync(join(out,'render.pdf'),Buffer.from(process.env.AGENT_DOCX_FAKE_PDF,'base64'));`;
  try {
    await writeFile(executable, script);
    await chmod(executable, 0o755);
    const result = await measureMarkdown("# Exact bytes", {
      renderer: "libreoffice",
      sectionDiagnostics: true,
      includeGeneratedDocx: true,
      libreoffice: { executablePath: executable },
    });
    const hash = createHash("sha256")
      .update(result.generatedDocx)
      .digest("hex");
    assert.equal(result.pageCountSource, "libreoffice");
    assert.equal(result.deterministic.sections[1].source, "deterministic");
    assert.equal("sections" in result.renderers.libreoffice.value, false);
    assert.equal(hash, result.renderers.libreoffice.value.generatedDocxSha256);
  } finally {
    delete process.env.AGENT_DOCX_FAKE_PDF;
    await rm(root, { recursive: true, force: true });
  }
});

test("Word bridge validates and maps version-2 Unicode paragraph records", async () => {
  const root = await mkdtemp(join(tmpdir(), "agent-docx-fake-word-"));
  const executable = join(root, "powershell");
  const script = `#!/usr/bin/env node
let input='';for await(const chunk of process.stdin)input+=chunk;
const request=JSON.parse(input);
if(request.protocolVersion!==2||request.paragraphIds.length!==2)process.exit(9);
console.log(JSON.stringify({kind:'started',protocolVersion:2,hwnd:1}));
console.log(JSON.stringify({kind:'summary',protocolVersion:2,pageCount:2,totalBodyLines:3,bodyLinesByPage:[2,1],bodyLinesOnLastPage:1,version:'99',build:'1',activePrinter:'fake'}));
console.log(JSON.stringify({kind:'paragraphs',protocolVersion:2,status:'ok',value:[
  {id:request.paragraphIds[0],lineCount:1,startPage:1,endPage:1,finalLineText:'Unicode § 😀'},
  {id:request.paragraphIds[1],lineCount:2,startPage:1,endPage:2,finalLineText:'final 1'}
]}));`;
  const manifest = [
    {
      id: "adx_body_000000",
      index: 0,
      position: {
        start: { line: 1, column: 1, offset: 0 },
        end: { line: 1, column: 4, offset: 3 },
      },
      preview: "One",
    },
    {
      id: "adx_body_000001",
      index: 1,
      position: {
        start: { line: 3, column: 1, offset: 5 },
        end: { line: 3, column: 4, offset: 8 },
      },
      preview: "Two",
    },
  ];
  try {
    await writeFile(executable, script);
    await chmod(executable, 0o755);
    const rendered = await renderWord(
      new Uint8Array([1, 2, 3]),
      [],
      { powerShellPath: executable },
      10000,
      manifest,
    );
    assert.equal(rendered.pageCount, 2);
    assert.deepEqual(rendered.paragraphDiagnostics, {
      status: "ok",
      value: [
        {
          source: "word",
          index: 0,
          position: manifest[0].position,
          startPage: 1,
          endPage: 1,
          lineCount: 1,
          finalLineText: "Unicode § 😀",
          preview: "One",
        },
        {
          source: "word",
          index: 1,
          position: manifest[1].position,
          startPage: 1,
          endPage: 2,
          lineCount: 2,
          finalLineText: "final 1",
          preview: "Two",
        },
      ],
    });
  } finally {
    await rm(root, { recursive: true, force: true });
  }
});

test("Word paragraph extraction failure preserves a successful summary", async () => {
  const root = await mkdtemp(join(tmpdir(), "agent-docx-fake-word-error-"));
  const executable = join(root, "powershell");
  const script = `#!/usr/bin/env node
for await(const chunk of process.stdin){}
console.log(JSON.stringify({kind:'started',protocolVersion:2,hwnd:1}));
console.log(JSON.stringify({kind:'summary',protocolVersion:2,pageCount:1,totalBodyLines:1,bodyLinesByPage:[1],bodyLinesOnLastPage:1,version:'99',build:'1',activePrinter:'fake'}));
console.log(JSON.stringify({kind:'paragraphs',protocolVersion:2,status:'error',message:'injected extraction failure'}));`;
  try {
    await writeFile(executable, script);
    await chmod(executable, 0o755);
    const rendered = await renderWord(
      new Uint8Array([1]),
      [],
      { powerShellPath: executable },
      10000,
      [
        {
          id: "adx_body_000000",
          index: 0,
          position: {
            start: { line: 1, column: 1, offset: 0 },
            end: { line: 1, column: 2, offset: 1 },
          },
          preview: "x",
        },
      ],
    );
    assert.equal(rendered.pageCount, 1);
    assert.equal(rendered.paragraphDiagnostics.status, "error");
    assert.equal(
      rendered.paragraphDiagnostics.error.phase,
      "paragraph-diagnostics",
    );
    assert.match(
      rendered.paragraphDiagnostics.error.message,
      /injected extraction failure/,
    );
  } finally {
    await rm(root, { recursive: true, force: true });
  }
});

test("Word bridge rejects malformed frame ordering and framing", async () => {
  const root = await mkdtemp(join(tmpdir(), "agent-docx-fake-word-bad-"));
  const summary = JSON.stringify({
    kind: "summary",
    protocolVersion: 2,
    pageCount: 1,
    totalBodyLines: 1,
    bodyLinesByPage: [1],
    bodyLinesOnLastPage: 1,
    version: "99",
    build: "1",
    activePrinter: "fake",
  });
  const started = JSON.stringify({
    kind: "started",
    protocolVersion: 2,
    hwnd: 1,
  });
  try {
    for (const [name, lines] of [
      ["missing-start", [summary]],
      ["non-json-frame", [started, "not json", summary]],
      ["wrong-order", [summary, started]],
      [
        "invalid-start-version",
        [
          JSON.stringify({
            kind: "started",
            protocolVersion: 1,
            hwnd: null,
          }),
          summary,
        ],
      ],
    ]) {
      const executable = join(root, name);
      await writeFile(
        executable,
        `#!/usr/bin/env node
for await(const chunk of process.stdin){}
for(const line of ${JSON.stringify(lines)})console.log(line);`,
      );
      await chmod(executable, 0o755);
      await assert.rejects(
        () =>
          renderWord(
            new Uint8Array([1]),
            [],
            { powerShellPath: executable },
            10000,
          ),
        (error) =>
          error instanceof AgentDocxError &&
          error.code === "WORD_RENDER_FAILED",
      );
    }
  } finally {
    await rm(root, { recursive: true, force: true });
  }
});

test(
  "native Word returns stable bookmark paragraph diagnostics",
  { skip: process.env.AGENT_DOCX_WORD_SMOKE !== "1" },
  async () => {
    const first = `§ 1983 ${"constitutional remedy ".repeat(700)}[^1]`;
    const result = await measureMarkdown(
      `${first}\n\nSecond paragraph final.\n\nUnicode final §.\n\n[^1]: Footnote text.`,
      {
        renderer: "word",
        paragraphDiagnostics: true,
      },
    );
    assert.equal(result.renderers.word.status, "ok");
    const diagnostics = result.renderers.word.value.paragraphDiagnostics;
    assert.equal(diagnostics.status, "ok");
    assert.equal(diagnostics.value.length, 3);
    assert.deepEqual(
      diagnostics.value.map(({ index }) => index),
      [0, 1, 2],
    );
    assert.ok(diagnostics.value[0].lineCount > 1);
    assert.ok(diagnostics.value[0].endPage > diagnostics.value[0].startPage);
    assert.doesNotMatch(diagnostics.value[0].finalLineText, /\u0002|\r/);
    assert.match(diagnostics.value[0].finalLineText, /1$/);
    assert.equal(diagnostics.value[1].finalLineText, "Second paragraph final.");
    assert.equal(diagnostics.value[2].finalLineText, "Unicode final §.");
  },
);
