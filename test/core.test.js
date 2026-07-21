import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { Document, Packer, Paragraph } from "docx";
import { builtInProfiles, estimateMarkdown, inspectDocxTemplate, measureMarkdown, AgentDocxError } from "../dist/index.js";

test("root API and immutable profiles",()=>{
  assert.deepEqual(Object.keys(builtInProfiles),["us-district-conventional","frap-32","cand-civil"]);
  assert.equal(Object.isFrozen(builtInProfiles),true);
  assert.equal(typeof estimateMarkdown,"function");
  assert.equal(typeof measureMarkdown,"function");
  assert.equal(typeof inspectDocxTemplate,"function");
});

test("empty Markdown has no deterministic physical page",async()=>{
  const result=await estimateMarkdown(" \n\t");
  assert.equal(result.pageCount,0);
  assert.equal(result.equivalentPages,0);
  assert.equal(result.lastPage,null);
});

test("exact 27-line boundary and 28th-line overflow",async()=>{
  const markdown27=await readFile(new URL("./fixtures/27-hard-lines.md",import.meta.url),"utf8");
  const markdown28=await readFile(new URL("./fixtures/28-hard-lines.md",import.meta.url),"utf8");
  const options={layout:{body:{lineSpacing:{rule:"exact",twips:480},beforeTwips:0,afterTwips:0,keepLines:false},pagination:{widowLines:1,orphanLines:1}}};
  const one=await estimateMarkdown(markdown27,options);
  const two=await estimateMarkdown(markdown28,options);
  assert.equal(one.pageCount,1);
  assert.equal(two.pageCount,2);
  assert.equal(two.equivalentPages,1+480/12960);
  assert.equal(two.lastPage.bodyLineEquivalentsUsed,1);
  assert.equal(two.lastPage.bodyLineCapacity,27);
});

test("explicit page break abandons remaining page",async()=>{
  const result=await estimateMarkdown("First.\n\n<!-- pagebreak -->\n\nSecond.");
  assert.equal(result.pageCount,2);
  assert.ok(result.equivalentPages>1&&result.equivalentPages<2);
});

test("unsupported Markdown rejects with source-aware code",async()=>{
  await assert.rejects(()=>estimateMarkdown("| A |\n| - |\n| x |"),error=>error instanceof AgentDocxError&&error.code==="UNSUPPORTED_MARKDOWN");
  await assert.rejects(()=>estimateMarkdown("```js\nalert(1)\n```"),error=>error instanceof AgentDocxError&&error.code==="UNSUPPORTED_MARKDOWN");
});

test("font shaping distinguishes narrow and wide glyphs",async()=>{
  const narrow=await estimateMarkdown("iiii",{paragraphDiagnostics:true});
  const wide=await estimateMarkdown("WWWW",{paragraphDiagnostics:true});
  assert.ok(wide.paragraphs[0].lastLineUsedTwips>narrow.paragraphs[0].lastLineUsedTwips);
});

test("CAND independent 28-counted-line ceiling",async()=>{
  const hard=n=>Array.from({length:n},(_,i)=>`x${i<n-1?"  ":""}`).join("\n");
  const layout={body:{lineSpacing:{rule:"exact",twips:400},keepLines:false},pagination:{widowLines:1,orphanLines:1}};
  assert.equal((await estimateMarkdown(hard(28),{profile:"cand-civil",layout})).pageCount,1);
  assert.equal((await estimateMarkdown(hard(29),{profile:"cand-civil",layout})).pageCount,2);
});

test("trim diagnostics are deterministic and advisory",async()=>{
  const result=await estimateMarkdown("Word ".repeat(250),{paragraphDiagnostics:true,trim:{maxLastLineRatio:1,maxCandidates:10},pageLimit:1});
  assert.ok(result.paragraphs.length===1);
  const p=result.paragraphs[0];
  assert.equal(p.lastLineRatio,p.lastLineUsedTwips/p.lastLineAvailableTwips);
  assert.equal(result.trimOpportunities[0].message,"Shortening or rephrasing this paragraph may remove its final wrapped line.");
  assert.equal(result.budget.limitPages,1);
});

test("generated DOCX inspection imports section geometry",async()=>{
  const doc=new Document({sections:[{properties:{page:{size:{width:11907,height:16839},margin:{top:1000,right:1100,bottom:1200,left:1300,header:500,footer:600,gutter:0}}},children:[new Paragraph("Template")]}]});
  const bytes=await Packer.toBuffer(doc);
  const inspected=await inspectDocxTemplate(bytes);
  assert.equal(inspected.package.macroEnabled,false);
  assert.equal(inspected.sections.at(-1).page.widthTwips,11907);
  assert.equal(inspected.imported.page.marginsTwips.left,1300);
});
