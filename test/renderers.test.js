import test from "node:test";
import assert from "node:assert/strict";
import { chmod, mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { PDFDocument } from "pdf-lib";
import { renderLibreOffice, resolveLibreOffice } from "../dist/renderers/office.js";
import { AgentDocxError } from "../dist/index.js";

test("LibreOffice resolver rejects an explicit missing executable",async()=>{
  await assert.rejects(()=>resolveLibreOffice("/definitely/missing/soffice"),error=>error instanceof AgentDocxError&&error.code==="LIBREOFFICE_NOT_FOUND");
});

test("LibreOffice adapter uses isolated exact conversion arguments and PDF page tree",async()=>{
  const root=await mkdtemp(join(tmpdir(),"agent-docx-fake-lo-"));
  const executable=join(root,"soffice");
  const pdf=await PDFDocument.create();pdf.addPage();pdf.addPage();
  process.env.AGENT_DOCX_FAKE_PDF=Buffer.from(await pdf.save()).toString("base64");
  const script=`#!/usr/bin/env node\nimport {writeFileSync} from 'node:fs';import {join} from 'node:path';const a=process.argv.slice(2);if(a[0]==='--version'){console.log('LibreOffice 99.0 fake');process.exit(0)}const expected=['--headless','--nologo','--nodefault','--norestore','--infilter=Office Open XML Text','--convert-to','pdf:writer_pdf_Export'];for(const x of expected)if(!a.includes(x)){console.error('missing '+x);process.exit(9)}const out=a[a.indexOf('--outdir')+1];writeFileSync(join(out,'render.pdf'),Buffer.from(process.env.AGENT_DOCX_FAKE_PDF,'base64'));`;
  try{
    await writeFile(executable,script);await chmod(executable,0o755);
    const rendered=await renderLibreOffice(new Uint8Array([1,2,3]),["Times New Roman"],{executablePath:executable},10000);
    assert.equal(rendered.pageCount,2);
    assert.match(rendered.versionRaw,/99.0 fake/);
    assert.equal(rendered.calibratedFontEnvironment,false);
    assert.equal(rendered.requestedFontFamilies[0],"Times New Roman");
  }finally{delete process.env.AGENT_DOCX_FAKE_PDF;await rm(root,{recursive:true,force:true})}
});
