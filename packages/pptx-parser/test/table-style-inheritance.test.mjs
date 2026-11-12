import assert from "node:assert/strict";
import test from "node:test";
import { normalizePresentation } from "@pptkit/core";
import { importPptx } from "../dist/index.js";
import { table as tableMarkup, tableRow, cell, tableFixture, tableTransform } from "./table-fixture.mjs";
import { encodeStructure } from "./structure-fixture.mjs";
import { tableStyleFixture, tableStyleBytes, TABLE_STYLE_ID, paint, line, region } from "./table-style-fixture.mjs";
import { readGeneratedZip } from "../../../tests/support/read-generated-zip.mjs";
const open = bytes => importPptx(bytes,{projection:"content",editSaveProfile:"text-runs-v3"});
const table = s => normalizePresentation(s.document).slides[0].elements[0];
const colors = s => table(s).rows.map(r=>r.cells.map(c=>c.style.fill.color));
const report = s => s.report.tableStyleInheritance;

test("both dialects resolve named/inline/default style flags, all region priorities and local explicit fill",async()=>{
  for(const strict of [false,true]) for(const mode of ["reference","inline","default"]){
    const session=await open(tableStyleBytes({strict,inline:mode==="inline",reference:mode!=="default",direct:`<a:tcPr>${paint("ABCDEF")}</a:tcPr>`}));
    try{
      assert.deepEqual(colors(session),[["990000","778899","778899","009900"],["556677","ABCDEF","445566","667788"],["556677","334455","445566","667788"],["000099","8899AA","8899AA","999900"]]);
      assert.equal(report(session).tables[0].status,mode==="inline"?"inline":"resolved"); assert.equal(report(session).tables[0].styleId,TABLE_STYLE_ID); assert.ok(report(session).tables[0].definition.path); assert.ok(report(session).properties.some(p=>p.region==="nwCell"&&p.property==="cell.fill"));
      const direct=report(session).properties.find(p=>p.property==="cell.fill"&&p.rowIndex===1&&p.columnIndex===1); assert.equal(direct.source.part,"/deck/z.xml"); assert.equal(direct.value.color,"ABCDEF");
      assert.equal(table(session).rows[2].cells[1].content[0].runs[0].style.bold,true); assert.ok(Object.isFrozen(report(session).properties[0]));
    }finally{session.dispose();}
  }
});

test("theme palette/maps/overrides and Latin fonts resolve in the target slide scope with complete dependencies",async()=>{
  const rules=`<a:wholeTbl><a:tcTxStyle b="off" i="on"><a:fontRef idx="minor"/><a:schemeClr val="accent1"><a:alpha val="50000"/></a:schemeClr></a:tcTxStyle><a:tcStyle><a:fill>${paint("accent1",true,'<a:alpha val="25000"/>')}</a:fill></a:tcStyle></a:wholeTbl>`;
  const f=tableStyleFixture({flags:"",rules}); f.parts['deck/master.xml']=f.parts['deck/master.xml'].replace('</p:sldMaster>','<p:clrMap accent1="accent2"/></p:sldMaster>');
  const session=await open(encodeStructure(f));
  try{
    const c=table(session).rows[0].cells[0],r=c.content[0].runs[0].style; assert.deepEqual(c.style.fill,{type:"solid",color:"CA6420",opacity:.25});assert.equal(r.color,"CA6420");assert.equal(r.colorOpacity,.5);assert.equal(r.fontFamily,"Arial");assert.equal(r.bold,false);assert.equal(r.italic,true);
    const property=report(session).properties.find(p=>p.property==="cell.fill");assert.ok(property.resolvedFrom.some(s=>s.part==="/deck/theme.xml"));assert.ok(property.resolvedFrom.some(s=>s.part==="/deck/master.xml")); assert.ok(!property.resolvedFrom.some(s=>s.path?.includes("undefined")));
  }finally{session.dispose();}
  const g=tableStyleFixture({flags:"",rules});g.parts['deck/_rels/z.xml.rels']=g.parts['deck/_rels/z.xml.rels'].replace('</Relationships>',`<Relationship Id="override" Type="${g.R}/themeOverride" Target="override.xml"/></Relationships>`);g.parts['deck/override.xml']=`<a:themeOverride xmlns:a="${g.A}"><a:clrScheme name="Local"><a:accent1><a:srgbClr val="EE2211"/></a:accent1></a:clrScheme></a:themeOverride>`;g.types['deck/override.xml']='themeOverride';
  const overridden=await open(encodeStructure(g));try{assert.equal(colors(overridden)[0][0],"EE2211");assert.ok(report(overridden).properties.some(p=>p.resolvedFrom?.some(s=>s.part==="/deck/override.xml")));}finally{overridden.dispose();}
});

test("whole/region physical outer and inner borders are projected without inventing merge interiors",async()=>{
  const rules=region("wholeTbl",paint("FFFFFF"),"",["left","right","top","bottom","insideH","insideV"].map((side,i)=>`<a:${side}>${line(["AA0000","00AA00","0000AA","AAAA00","AA00AA","00AAAA"][i],i+1)}</a:${side}>`).join(""))+region("firstRow","","",`<a:bottom>${line("123456",7)}</a:bottom>`);
  const session=await open(tableStyleBytes({rules,flags:'firstRow="1"'}));
  try{
    const rows=table(session).rows;assert.equal(rows[0].cells[0].style.borders.left.paint.color,"AA0000");assert.equal(rows[0].cells[1].style.borders.left.paint.color,"00AAAA");assert.equal(rows[0].cells[1].style.borders.bottom.paint.color,"123456");assert.equal(rows[1].cells[1].style.borders.top.paint.color,"AA00AA");assert.equal(rows.at(-1).cells[0].style.borders.bottom.paint.color,"AAAA00");
    assert.ok(report(session).properties.some(p=>p.property==="cell.border.left"&&p.region==="wholeTbl"));
  }finally{session.dispose();}
});

test("unsupported high-priority paint/font/line declarations block older values and keep explicit findings",async()=>{
  const rules=region("wholeTbl",paint("FF0000"))+`<a:firstRow><a:tcTxStyle b="def"><a:fontRef idx="none"/><a:prstClr val="red"/></a:tcTxStyle><a:tcStyle><a:fill><a:gradFill/></a:fill><a:tcBdr><a:left><a:lnRef idx="1"/></a:left></a:tcBdr></a:tcStyle></a:firstRow>`;
  const session=await open(tableStyleBytes({flags:'firstRow="1"',rules}));
  try{
    assert.equal(table(session).rows[0].cells[0].style.fill.type,"none"); assert.equal(colors(session)[1][0],"FF0000");assert.ok(report(session).properties.some(p=>p.status==="unsupported"&&p.property==="cell.fill"));assert.ok(report(session).properties.some(p=>p.status==="unsupported"&&p.property==="run.fontFamily"));assert.ok(session.report.findings.some(f=>/format scheme|style-matrix|fillRef|lnRef/.test(f.reason)));
  }finally{session.dispose();}
});

test("missing and ambiguous identities, invalid relationship/type/flags and MCE alternatives never silently choose a style",async()=>{
  const missing=tableStyleFixture();missing.parts['deck/z.xml']=missing.parts['deck/z.xml'].replace(TABLE_STYLE_ID,'{00000000-0000-0000-0000-000000000000}');const s=await open(encodeStructure(missing));try{assert.equal(report(s).tables[0].status,"missing");assert.equal(table(s).rows[0].cells[0].style.fill.type,"none");}finally{s.dispose();}
  for(const mutate of [
    f=>f.parts['deck/tableStyles.xml']=f.parts['deck/tableStyles.xml'].replace('</a:tblStyleLst>',`<a:tblStyle styleId="${TABLE_STYLE_ID.toLowerCase()}"/></a:tblStyleLst>`),
    f=>f.parts['deck/_rels/main.data.rels']=f.parts['deck/_rels/main.data.rels'].replace('Target="tableStyles.xml"','Target="https://example.invalid/styles.xml" TargetMode="External"'),
    f=>f.types['deck/tableStyles.xml']='theme',
    f=>f.parts['deck/z.xml']=f.parts['deck/z.xml'].replace('firstRow="1"','firstRow="maybe"'),
  ]){const f=tableStyleFixture();mutate(f);await assert.rejects(open(encodeStructure(f)));}
  const m=tableStyleFixture();m.parts['deck/tableStyles.xml']=m.parts['deck/tableStyles.xml'].replace(`def="${TABLE_STYLE_ID}"`,`def="${TABLE_STYLE_ID}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:unknown"`).replace(/<a:wholeTbl>[\s\S]*?<\/a:wholeTbl>/,old=>'<mc:AlternateContent><mc:Choice Requires="x"><a:wholeTbl><a:tcStyle><a:fill>'+paint('FF00FF')+'</a:fill></a:tcStyle></a:wholeTbl></mc:Choice><mc:Fallback>'+old+'</mc:Fallback></mc:AlternateContent>');
  const compatible=await open(encodeStructure(m));try{assert.equal(report(compatible).tables[0].status,"resolved");assert.ok(compatible.report.inspection.markupCompatibility.some(x=>x.part==="/deck/tableStyles.xml"));}finally{compatible.dispose();}
});

test("three scalar saves retain style/theme bytes and unsupported style mutations block atomically with restore retry",async()=>{
  let bytes=tableStyleBytes({flags:""});
  for(let round=0;round<3;round++){
    const session=await open(bytes);try{
      const owner=session.document.slides[0],base=owner.elements[0],changed=structuredClone(base);changed.rows[1].cells[1].content[0].runs[0].text=`Round ${round}`;owner.replaceElement(base.id,changed);
      const plan=await session.prepareSave({mode:"edit"});assert.equal(plan.status,"ready");const saved=await session.save(plan.plan),before=readGeneratedZip(bytes),after=readGeneratedZip(saved.bytes);for(const [name,value]of before)if(name!=="deck/z.xml")assert.deepEqual(after.get(name),value);
      const bad=structuredClone(changed);bad.rows[0].cells[0].style.fill={type:"none"};owner.replaceElement(base.id,bad);assert.equal((await session.prepareSave({mode:"edit"})).status,"blocked");owner.replaceElement(base.id,changed);assert.equal((await session.prepareSave({mode:"edit"})).status,"ready");bytes=saved.bytes;
    }finally{session.dispose();}
  }
});

test("table provenance/work/dependency limits have exact session-wide thresholds and failed attempts can retry",async()=>{
  const bytes=tableStyleBytes({flags:"",rules:region("wholeTbl",paint("accent1",true))});
  const baseline=await open(bytes);const expected=normalizePresentation(baseline.document);baseline.dispose();
  for(const key of ["maxTableStyleProperties","maxTableStyleDependencies","maxTableStyleWork"]){
    let lo=1,hi=2000000;while(lo<hi){const mid=Math.floor((lo+hi)/2);try{const s=await importPptx(bytes,{projection:"content",sessionLimits:{[key]:mid}});s.dispose();hi=mid;}catch(e){assert.equal(e.code,"import-limit-exceeded");lo=mid+1;}}
    const s=await importPptx(bytes,{projection:"content",sessionLimits:{[key]:lo}});assert.deepEqual(normalizePresentation(s.document),expected);s.dispose();if(lo>1)await assert.rejects(importPptx(bytes,{projection:"content",sessionLimits:{[key]:lo-1}}),{code:"import-limit-exceeded"});
  }
});

test("one shared style definition resolves independently under two masters without cache contamination",async()=>{
  const f=tableStyleFixture({flags:"",rules:region("wholeTbl",paint("accent1",true))});
  f.parts['deck/a.xml']=f.parts['deck/z.xml'];f.parts['deck/_rels/a.xml.rels']=f.parts['deck/_rels/a.xml.rels'].replace('layout.xml','layout2.xml');
  f.parts['deck/layout2.xml']=f.parts['deck/layout.xml'];f.parts['deck/_rels/layout2.xml.rels']=f.parts['deck/_rels/layout.xml.rels'].replace('master.xml','master2.xml');
  f.parts['deck/master2.xml']=f.parts['deck/master.xml'];f.parts['deck/_rels/master2.xml.rels']=f.parts['deck/_rels/master.xml.rels'].replace('layout.xml','layout2.xml').replace('theme.xml','theme2.xml');f.parts['deck/theme2.xml']=f.parts['deck/theme.xml'].replace('3579BD','1122FF');
  f.parts['deck/main.data']=f.parts['deck/main.data'].replace('</p:sldMasterIdLst>','<p:sldMasterId id="2147483650" r:id="m2"/></p:sldMasterIdLst>');
  f.parts['deck/_rels/main.data.rels']=f.parts['deck/_rels/main.data.rels'].replace('</Relationships>',`<Relationship Id="m2" Type="${f.R}/slideMaster" Target="master2.xml"/></Relationships>`);
  Object.assign(f.types,{'deck/layout2.xml':'presentationml.slideLayout','deck/master2.xml':'presentationml.slideMaster','deck/theme2.xml':'theme'});
  const session=await open(encodeStructure(f));try{
    const p=normalizePresentation(session.document);assert.equal(p.slides[0].elements[0].rows[0].cells[0].style.fill.color,"3579BD");assert.equal(p.slides[1].elements[0].rows[0].cells[0].style.fill.color,"1122FF");
    assert.equal(report(session).tables.length,2);assert.ok(report(session).properties.some(p=>p.target.part==="/deck/a.xml"&&p.resolvedFrom.some(s=>s.part==="/deck/theme2.xml")));
    const bound=report(session).properties.length+report(session).tables.length;await assert.rejects(importPptx(encodeStructure(f),{projection:"content",sessionLimits:{maxTableStyleProperties:bound-1}}),{code:"import-limit-exceeded"});
  }finally{session.dispose();}
  f.parts['deck/theme.xml']=f.parts['deck/theme.xml'].replace(/<a:accent1>[\s\S]*?<\/a:accent1>/,'<a:accent1><a:prstClr val="red"/></a:accent1>');
  const unsupported=await open(encodeStructure(f));try{
    const p=normalizePresentation(unsupported.document);assert.equal(p.slides[0].elements[0].rows[0].cells[0].style.fill.type,"none");assert.equal(p.slides[1].elements[0].rows[0].cells[0].style.fill.color,"1122FF");
    assert.equal(unsupported.report.findings.filter(f=>f.feature==="drawing.table-style.unmapped-property"&&/Theme slot accent1/.test(f.reason)).length,1,"one unsupported source scope must not taint the valid scope's audit");
  }finally{unsupported.dispose();}
});

test("pre-aborted/during-import cancellation and async option/byte capture leave independent retries usable",async()=>{
  const bytes=tableStyleBytes(),controller=new AbortController();controller.abort();await assert.rejects(importPptx(bytes,{projection:"content",signal:controller.signal}),{code:"operation-aborted"});
  const live=new AbortController(),pending=importPptx(bytes,{projection:"content",signal:live.signal});queueMicrotask(()=>live.abort());await assert.rejects(pending,{code:"operation-aborted"});
  const source=Uint8Array.from(bytes),options={projection:"content",sessionLimits:{maxTableStyleWork:2000000}};const captured=importPptx(source,options);source.fill(0);options.sessionLimits.maxTableStyleWork=1;
  const result=await captured;try{assert.equal(report(result).tables[0].status,"resolved");}finally{result.dispose();}
  const retry=await open(bytes);retry.dispose();
});


test("heterogeneous merged physical borders block inherited sides and source scalar saves preserve every conditional style byte",async()=>{
  const f=tableStyleFixture({flags:'firstRow="1"',rules:region("wholeTbl",paint("FFFFFF"),"",`<a:insideV>${line("777777",2)}</a:insideV>`)+region("firstRow",paint("334455"))});
  const p=border=>`<a:tcPr>${border}</a:tcPr>`,right=color=>line(color,3).replace('<a:ln ','<a:lnR ').replace('</a:ln>','</a:lnR>');
  const rows=[
    tableRow([cell({text:"Merged",attributes:'gridSpan="2" rowSpan="2"'}),cell({attributes:'hMerge="1" rowSpan="2"',properties:p(right("00AA00"))}),cell({text:"Upper"}),cell({text:"Control A"})],30),
    tableRow([cell({attributes:'vMerge="1" gridSpan="2"'}),cell({attributes:'hMerge="1" vMerge="1"',properties:p(right("0000AA"))}),cell({text:"Lower"}),cell({text:"Control B"})],30),
    tableRow(Array.from({length:4},()=>cell({text:"Footer"})),30),
  ];
  const markup=tableMarkup({columns:[70,70,70,70],rows,transform:tableTransform("",280,90),properties:`<a:tblPr firstRow="1"><a:tableStyleId>${TABLE_STYLE_ID}</a:tableStyleId></a:tblPr>`});
  f.parts['deck/z.xml']=tableFixture(markup).parts['deck/z.xml'];const bytes=encodeStructure(f),session=await open(bytes);
  try{
    assert.equal(table(session).rows[0].cells[0].style.borders.right.paint.type,"none");assert.equal(table(session).rows[0].cells[0].style.borders.right.width,0);
    assert.ok(report(session).properties.some(p=>p.property==="cell.border.right"&&p.rowIndex===0&&p.columnIndex===0&&p.status==="unsupported"));assert.ok(session.report.findings.some(f=>f.operation==="preview"&&f.feature==="drawing.table.mapped-fields"&&f.tier==="fallback"));
    const slide=session.document.slides[0],base=slide.elements[0],changed=structuredClone(base);changed.rows[0].cells[0].content[0].runs[0].text="Edited merged";slide.replaceElement(base.id,changed);const plan=await session.prepareSave({mode:"edit"});assert.equal(plan.status,"ready");const saved=readGeneratedZip((await session.save(plan.plan)).bytes);assert.deepEqual(saved.get("deck/tableStyles.xml"),readGeneratedZip(bytes).get("deck/tableStyles.xml"));
  }finally{session.dispose();}
});

test("explicit table definitions win even when their IDs match a built-in preset", async () => {
  for (const strict of [false, true]) for (const inline of [false, true]) {
    const fixture = tableStyleFixture({ strict, inline });
    const preset = "{21E4AEA4-8DFA-4A89-87EB-49C32662AFE0}";
    for (const key of Object.keys(fixture.parts)) {
      if (typeof fixture.parts[key] === "string") fixture.parts[key] = fixture.parts[key].replaceAll(TABLE_STYLE_ID, preset);
    }
    const bytes = encodeStructure(fixture), session = await open(bytes);
    try {
      assert.equal(table(session).nativeTableStylePreset, undefined);
      assert.equal(report(session).tables[0].status, inline ? "inline" : "resolved");
      assert.equal(report(session).tables[0].styleId, preset);
      assert.equal(colors(session)[0][0], "990000");
      const plan = await session.prepareSave({ mode: "no-op" });
      assert.equal(plan.status, "ready");
      assert.deepEqual((await session.save(plan.plan)).bytes, new Uint8Array(bytes));
    } finally { session.dispose(); }
  }
});

test("a compatibility-projected registry cannot establish that a built-in ID has no definition", async () => {
  for (const strict of [false, true]) {
    const fixture = tableStyleFixture({ strict });
    const preset = "{21E4AEA4-8DFA-4A89-87EB-49C32662AFE0}";
    fixture.parts["deck/z.xml"] = fixture.parts["deck/z.xml"].replaceAll(TABLE_STYLE_ID, preset);
    const original = fixture.parts["deck/tableStyles.xml"];
    const definition = original.match(/<a:tblStyle\b[\s\S]*?<\/a:tblStyle>/)[0].replaceAll(TABLE_STYLE_ID, preset);
    fixture.parts["deck/tableStyles.xml"] = `<a:tblStyleLst xmlns:a="${fixture.A}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:unsupported" def="${preset}"><mc:AlternateContent><mc:Choice Requires="x">${definition}</mc:Choice><mc:Fallback/></mc:AlternateContent></a:tblStyleLst>`;
    const bytes = encodeStructure(fixture), session = await open(bytes);
    try {
      assert.equal(table(session).nativeTableStylePreset, undefined);
      const plan = await session.prepareSave({ mode: "no-op" });
      assert.equal(plan.status, "ready");
      assert.deepEqual((await session.save(plan.plan)).bytes, new Uint8Array(bytes));
    } finally { session.dispose(); }
  }
});

test("unrelated compatibility branches do not hide a bare built-in table preset", async () => {
  for (const strict of [false, true]) {
    const fixture = tableStyleFixture({ strict });
    const preset = "{21E4AEA4-8DFA-4A89-87EB-49C32662AFE0}";
    fixture.parts["deck/z.xml"] = fixture.parts["deck/z.xml"].replaceAll(TABLE_STYLE_ID, preset);
    const original = fixture.parts["deck/tableStyles.xml"];
    const definition = original.match(/<a:tblStyle\b[\s\S]*?<\/a:tblStyle>/)[0];
    fixture.parts["deck/tableStyles.xml"] = `<a:tblStyleLst xmlns:a="${fixture.A}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:unsupported"><mc:AlternateContent><mc:Choice Requires="x">${definition}</mc:Choice><mc:Fallback/></mc:AlternateContent></a:tblStyleLst>`;
    const bytes = encodeStructure(fixture), session = await open(bytes);
    try {
      assert.equal(table(session).nativeTableStylePreset, "medium-2-accent-2");
      const plan = await session.prepareSave({ mode: "no-op" });
      assert.equal(plan.status, "ready");
      assert.deepEqual((await session.save(plan.plan)).bytes, new Uint8Array(bytes));
    } finally { session.dispose(); }
  }
});

test("raw registry scan is charged only when a known bare preset needs collision proof", async () => {
  const fixture = tableStyleFixture();
  const base = fixture.parts["deck/z.xml"];
  async function minimum(id) {
    fixture.parts["deck/z.xml"] = base.replaceAll(TABLE_STYLE_ID, id);
    const bytes = encodeStructure(fixture);
    let low = 1, high = 10000;
    while (low < high) {
      const limit = Math.floor((low + high) / 2);
      try {
        const session = await importPptx(bytes, { projection: "content", sessionLimits: { maxTableStyleWork: limit } });
        session.dispose();
        high = limit;
      } catch (error) {
        assert.equal(error.code, "import-limit-exceeded");
        low = limit + 1;
      }
    }
    return low;
  }
  const unknown = await minimum("{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}");
  const known = await minimum("{21E4AEA4-8DFA-4A89-87EB-49C32662AFE0}");
  assert.ok(known > unknown, `collision proof ${known} must cost more than unrelated reference ${unknown}`);
});
