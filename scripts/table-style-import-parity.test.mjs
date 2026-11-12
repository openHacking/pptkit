import assert from "node:assert/strict";
import test from "node:test";
import { createHash } from "node:crypto";
import { readFile } from "node:fs/promises";
import { chromium } from "@playwright/test";
import { normalizePresentation } from "../packages/core/dist/index.js";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { tableStyleBytes, tableStyleFixture, paint, region } from "../packages/pptx-parser/test/table-style-fixture.mjs";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";
const profile={projection:"content",editSaveProfile:"text-runs-v3"};
const moduleUrl=text=>`data:text/javascript;base64,${Buffer.from(text).toString("base64")}`;
async function rounds(bytes){
  const results=[];
  for(let round=0;round<3;round++){
    const session=await importPptx(bytes,profile);try{
      const document=normalizePresentation(session.document),report=session.report.tableStyleInheritance,slide=session.document.slides[0],old=slide.elements[0],changed=structuredClone(old);changed.rows[1].cells[1].content[0].runs[0].text=`Saved ${round}`;slide.replaceElement(old.id,changed);
      const ready=await session.prepareSave({mode:"edit"});assert.equal(ready.status,"ready");const saved=await session.save(ready.plan);results.push({document,report,saved:[...saved.bytes]});bytes=saved.bytes;
    }finally{session.dispose();}
  }
  return results;
}
test("Node/ESM/global table-style provenance and three source saves agree in both dialects, with independent DOM source resolution",async t=>{
  const browser=await chromium.launch();t.after(()=>browser.close());
  const core=moduleUrl(await readFile("packages/core/dist/browser/index.js","utf8")),parser=moduleUrl((await readFile("packages/pptx-parser/dist/browser/index.js","utf8")).replaceAll('"@pptkit/core"',JSON.stringify(core)));
  for(const strict of [false,true]){
    const bytes=tableStyleBytes({strict}),expected=await rounds(bytes),fixture=tableStyleFixture({strict});
    for(const format of ["esm","global"]){
      const page=await browser.newPage(),requests=[];await page.route("**/*",route=>{requests.push(route.request().url());return route.abort();});
      if(format==="global")for(const name of ["core","pptx-parser"])await page.addScriptTag({content:await readFile(`packages/${name}/dist/browser/global.js`,"utf8")});
      const result=await page.evaluate(async({core,parser,bytes,parts,namespace})=>{
        const c=core?await import(core):PPTKit,p=parser?await import(parser):PPTKit,results=[];
        for(let round=0;round<3;round++){
          const session=await p.importPptx(Uint8Array.from(bytes),{projection:"content",editSaveProfile:"text-runs-v3"});
          const document=c.normalizePresentation(session.document),report=session.report.tableStyleInheritance,slide=session.document.slides[0],old=slide.elements[0],changed=structuredClone(old);changed.rows[1].cells[1].content[0].runs[0].text=`Saved ${round}`;slide.replaceElement(old.id,changed);
          const ready=await session.prepareSave({mode:"edit"});if(ready.status!=="ready")throw new Error("Not ready");const saved=await session.save(ready.plan);results.push({document,report,saved:[...saved.bytes]});bytes=[...saved.bytes];session.dispose();
        }
        const docs=Object.fromEntries(Object.entries(parts).filter(([,v])=>typeof v==="string"&&v.startsWith("<")).map(([part,xml])=>[`/${part}`,new DOMParser().parseFromString(xml,"application/xml")]));
        let verified=0;
        for(const info of results[0].report.properties){
          if(info.property!=="cell.fill"||info.status!=="mapped"||info.value.type!=="solid")continue;
          const doc=docs[info.source.part],source=doc.evaluate(info.source.path,doc,prefix=>prefix==="a"?namespace:null,XPathResult.FIRST_ORDERED_NODE_TYPE,null).singleNodeValue;
          if(source?.localName!=="srgbClr"||source.getAttribute("val")!==info.value.color)throw new Error("Independent source/property mismatch");verified++;
        }
        return {results,verified};
      },{core:format==="esm"?core:null,parser:format==="esm"?parser:null,bytes:[...bytes],parts:fixture.parts,namespace:fixture.A});
      assert.deepEqual(result.results,expected);assert.equal(result.verified,16);assert.deepEqual(requests,[]);await page.close();
    }
    const before=readGeneratedZip(bytes),after=readGeneratedZip(Uint8Array.from(expected.at(-1).saved));for(const[name,value]of before)if(name!=="deck/z.xml")assert.deepEqual(after.get(name),value);
  }
});
const effective=document=>{
  const normalized=normalizePresentation(document),theme=normalized.theme;
  return normalized.slides[0].elements[0].rows.map(row=>row.cells.map(c=>({fill:c.style.fill,spans:[c.rowSpan,c.colSpan],text:c.content.map(p=>({style:p.style,runs:p.runs.map(r=>({...r,style:{...r.style,fontFamily:typeof r.style.fontFamily==="string"?r.style.fontFamily:theme.fonts[r.style.fontFamily.theme],color:typeof r.style.color==="string"?r.style.color:theme.colors[r.style.color.theme]}}))}))})));
};
test("fresh materialization retains effective source style formatting while table-style definitions and style mutations remain separate",async()=>{
  let bytes=tableStyleBytes({flags:"",rules:region("wholeTbl",paint("3579BD"))});let expected;
  for(let round=0;round<3;round++){
    const s=await importPptx(bytes,profile);try{
      expected??=effective(s.document);assert.deepEqual(effective(s.document),expected);
      const slide=s.document.slides[0],base=slide.elements[0],changed=structuredClone(base);changed.rows[0].cells[0].style.fill={type:"solid",color:"ABCDEF"};slide.replaceElement(base.id,changed);assert.equal((await s.prepareSave({mode:"edit"})).status,"blocked");slide.replaceElement(base.id,base);assert.equal((await s.prepareSave({mode:"no-op"})).status,"ready");
      bytes=(await generatePptx(s.document)).bytes;
    }finally{s.dispose();}
  }
});

test("PowerPoint native first-row toggles survive fresh round trips without becoming cell headers", async () => {
  const evidence = JSON.parse(await readFile("docs/testing/powerpoint-table-header-evidence.json", "utf8"));
  for (const [name, expected, record] of [
    ["table-header-off-powerpoint.pptx", false, evidence.headerOff],
    ["table-header-insert-powerpoint.pptx", true, evidence.headerInserted],
  ]) {
    let bytes = new Uint8Array(await readFile(`tests/fixtures/${name}`));
    assert.equal(createHash("sha256").update(bytes).digest("hex"), record.sha256);
    for (let round = 0; round <= 3; round++) {
      const session = await importPptx(bytes, profile);
      try {
        const slide = session.document.slides[0], table = slide.elements[0];
        assert.equal(table.type, "table");
        assert.equal(table.nativeFirstRowStyle, expected);
        assert.deepEqual(table.rows.map(row => row.cells.map(cell => cell.content.flatMap(p => p.runs.map(run => run.text)).join(""))), record.rows);
        assert.ok(table.rows.every(row => row.cells.every(cell => cell.headerScope === undefined)));
        const noOp = await session.prepareSave({ mode: "no-op" });
        assert.equal(noOp.status, "ready");
        assert.deepEqual((await session.save(noOp.plan)).bytes, bytes);
        if (round < 3) {
          bytes = (await generatePptx(session.document)).bytes;
          const xml = readGeneratedZip(bytes).get("ppt/slides/slide1.xml").toString();
          assert.match(xml, new RegExp(`<a:tblPr firstRow="${Number(expected)}" bandRow="1">`));
        }
        slide.replaceElement(table.id, { ...table, nativeFirstRowStyle: !expected });
        assert.notEqual((await session.prepareSave({ mode: "edit" })).status, "ready");
      } finally { session.dispose(); }
    }
  }
});

test("PowerPoint style switch matrix survives three fresh imports independently of cell headers", async () => {
  const evidence = JSON.parse(await readFile("docs/testing/powerpoint-table-style-flags-evidence.json", "utf8"));
  let bytes = new Uint8Array(await readFile(evidence.fixture));
  assert.equal(createHash("sha256").update(bytes).digest("hex"), evidence.savedSha256);
  const expected = { firstColumn: true, lastColumn: true, lastRow: true, bandedRows: false, bandedColumns: true };
  for (let round = 0; round <= 3; round++) {
    const session = await importPptx(bytes, profile);
    try {
      const slide = session.document.slides[0], table = slide.elements[0];
      assert.deepEqual(table.nativeTableStyleFlags, expected);
      assert.equal(table.nativeFirstRowStyle, false);
      assert.deepEqual(table.rows.map(row => row.cells.map(cell => cell.content.flatMap(p => p.runs.map(run => run.text)).join(""))), evidence.rows);
      assert.ok(table.rows.every(row => row.cells.every(cell => cell.headerScope === undefined)));
      const noOp = await session.prepareSave({ mode: "no-op" });
      assert.equal(noOp.status, "ready");
      assert.deepEqual((await session.save(noOp.plan)).bytes, bytes);
      if (round < 3) {
        bytes = (await generatePptx(session.document)).bytes;
        assert.match(readGeneratedZip(bytes).get("ppt/slides/slide1.xml").toString(),
          /<a:tblPr firstRow="0" bandRow="0" lastRow="1" firstCol="1" lastCol="1" bandCol="1">/);
      }
      slide.replaceElement(table.id, { ...table, nativeTableStyleFlags: { ...expected, lastRow: false } });
      assert.notEqual((await session.prepareSave({ mode: "edit" })).status, "ready");
    } finally { session.dispose(); }
  }
});

test("verified reference-only table preset round trips while a matching explicit definition is not replaced", async () => {
  const evidence = JSON.parse(await readFile("docs/testing/powerpoint-table-preset-evidence.json", "utf8"));
  for (const [record, expected] of [[evidence.referenceOnly, "medium-2-accent-2"], [evidence.defined, undefined]]) {
    let bytes = new Uint8Array(await readFile(record.fixture));
    assert.equal(createHash("sha256").update(bytes).digest("hex"), record.sha256);
    for (let round = 0; round <= (expected ? 3 : 0); round++) {
      const session = await importPptx(bytes, profile);
      try {
        const table = session.document.slides[0].elements[0];
        assert.equal(table.nativeTableStylePreset, expected);
        assert.ok(table.rows.every(row => row.cells.every(cell => cell.headerScope === undefined)));
        const plan = await session.prepareSave({ mode: "no-op" });
        assert.equal(plan.status, "ready");
        assert.deepEqual((await session.save(plan.plan)).bytes, bytes);
        if (expected && round < 3) {
          bytes = (await generatePptx(session.document)).bytes;
          const parts = readGeneratedZip(bytes);
          assert.ok(parts.get("ppt/slides/slide1.xml").toString().includes(`<a:tableStyleId>${record.reference}</a:tableStyleId>`));
          assert.doesNotMatch(parts.get("ppt/tableStyles.xml").toString(), /<a:tblStyle\b/);
        }
      } finally { session.dispose(); }
    }
  }
});
