import assert from "node:assert/strict";
import test from "node:test";
import { normalizePresentation } from "@pptkit/core";
import { importPptx } from "../dist/index.js";
import { encodeStructure } from "./structure-fixture.mjs";
import { matrixFixture,matrixOverride,formatScheme,matrixFill,matrixLine,fillReference,lineReference } from "./theme-matrix-fixture.mjs";
import { readGeneratedZip } from "../../../tests/support/read-generated-zip.mjs";
const open=f=>importPptx(encodeStructure(f),{projection:"content",editSaveProfile:"text-runs-v3"});
const cell=s=>normalizePresentation(s.document).slides[0].elements[0].rows[0].cells[0];
const properties=s=>s.report.tableStyleInheritance.properties;

test("Strict/Transitional fill and line matrix references resolve phClr, geometry and real source dependencies",async()=>{
  for(const strict of[false,true]){
    const s=await open(matrixFixture({strict}));try{
      assert.deepEqual(cell(s).style.fill,{type:"solid",color:"3579BD",opacity:1});assert.equal(cell(s).style.stroke.width,1);assert.equal(cell(s).style.stroke.paint.color,"CA6420");assert.equal(cell(s).style.stroke.dash,"solid");
      const p=properties(s).find(p=>p.property==="cell.fill");assert.ok(p.source.path.endsWith('/a:fillRef[1]/@idx'));assert.ok(p.resolvedFrom.some(p=>p.part==="/deck/theme.xml"&&/fmtScheme.*fillStyleLst.*solidFill/.test(p.path)));assert.ok(p.resolvedFrom.some(p=>p.part==="/deck/tableStyles.xml"&&/fillRef.*schemeClr/.test(p.path)));assert.ok(!properties(s).some(p=>JSON.stringify(p).includes("program")));
      assert.ok(!s.report.findings.some(f=>f.feature==="drawing.table-style.unmapped-property"&&/miter|cap|cmpd|algn/.test(f.source.path)));
    }finally{s.dispose();}
  }
});

test("fill 0/1000 are no-fill sentinels; 1-based normal/background slots and line zero remain distinct",async()=>{
  for(const[index,type,color]of[[0,"none"],[1000,"none"],[1,"solid","3579BD"],[2,"solid","123456"],[3,"none"],[1001,"solid","ABCDEF"],[1002,"solid","3579BD"]]){
    const s=await open(matrixFixture({fill:fillReference(index)}));try{assert.equal(cell(s).style.fill.type,type);if(color)assert.equal(cell(s).style.fill.color,color);if(index===1002)assert.equal(cell(s).style.fill.opacity,.25);}finally{s.dispose();}
  }
  const zero=await open(matrixFixture({scheme:"",fill:fillReference(0,'<a:schemeClr val="phClr"/>'),border:lineReference(0)}));try{assert.equal(cell(zero).style.fill.type,"none");assert.equal(cell(zero).style.stroke.paint.type,"none");assert.ok(properties(zero).filter(p=>p.property.startsWith("cell.")).every(p=>p.status==="mapped"));}finally{zero.dispose();}
  const line=await open(matrixFixture({border:lineReference(1001)}));try{assert.ok(properties(line).some(p=>p.property==="cell.border.left"&&p.status==="unsupported"));}finally{line.dispose();}
});

test("placeholder programs preserve ordered transforms and alpha reset/modulation without intermediate RGB rounding",async()=>{
  const ref='<a:srgbClr val="000000"><a:tint val="99500"/><a:alpha val="40000"/></a:srgbClr>';
  const f=matrixFixture({fill:fillReference(1,ref),scheme:formatScheme({fills:[matrixFill("phClr",'<a:shade val="50000"/><a:alphaMod val="50000"/>')]})});
  const s=await open(f);try{
    // Independent IEC sRGB single-channel oracle: linear blend then shade,
    // quantized once; an intermediate 8-bit quantization yields a different code.
    const linear=v=>v<=.04045?v/12.92:((v+.055)/1.055)**2.4,encoded=v=>v<=.0031308?v*12.92:1.055*v**(1/2.4)-.055;
    const exact=Math.round(encoded((1-(1-linear(0))*.995)*.5)*255),rounded=Math.round(encoded(linear(Math.round(encoded(1-(1-linear(0))*.995)*255)/255)*.5)*255);
    assert.notEqual(exact,rounded);assert.equal(cell(s).style.fill.color,exact.toString(16).padStart(2,"0").toUpperCase().repeat(3));assert.equal(cell(s).style.fill.opacity,.2);
  }finally{s.dispose();}
  const reset=await open(matrixFixture({fill:fillReference(1,'<a:schemeClr val="accent1"><a:alpha val="10000"/></a:schemeClr>'),scheme:formatScheme({fills:[matrixFill("phClr",'<a:alpha val="70000"/>')]})}));try{assert.equal(cell(reset).style.fill.opacity,.7);}finally{reset.dispose();}
});

test("fixed matrix RGB ignores unused placeholder color, while required missing/unbound/unknown colors block older declarations",async()=>{
  const fixed=await open(matrixFixture({fill:fillReference(2,'<a:prstClr val="unsupported"/>')}));try{assert.equal(cell(fixed).style.fill.color,"123456");}finally{fixed.dispose();}
  for(const ref of[fillReference(1,""),fillReference(1,'<a:schemeClr val="phClr"/>'),fillReference(1,'<a:schemeClr val="accent1"><a:future/></a:schemeClr>')]){
    const s=await open(matrixFixture({fill:ref}));try{assert.equal(cell(s).style.fill.type,"none");assert.ok(properties(s).some(p=>p.property==="cell.fill"&&p.status==="unsupported"));}finally{s.dispose();}
  }
  const unknown=await open(matrixFixture({scheme:formatScheme({fills:[matrixFill("phClr",'<a:future/>')]})}));try{assert.equal(cell(unknown).style.fill.type,"none");assert.ok(properties(unknown).find(p=>p.property==="cell.fill").resolvedFrom.some(p=>p.part==="/deck/theme.xml"));}finally{unknown.dispose();}
});

test("format-scheme and color components select independently; nearer partial/unsupported format schemes block the base",async()=>{
  const f=matrixOverride(matrixFixture(),formatScheme({fills:[matrixFill("phClr",'<a:alpha val="50000"/>')]}),'<a:clrScheme name="Local"><a:accent1><a:srgbClr val="ABCDEF"/></a:accent1><a:accent2><a:srgbClr val="001122"/></a:accent2></a:clrScheme>');
  const s=await open(f);try{assert.deepEqual(cell(s).style.fill,{type:"solid",color:"ABCDEF",opacity:.5});assert.ok(properties(s).find(p=>p.property==="cell.fill").resolvedFrom.some(p=>p.part==="/deck/matrix-override.xml"&&/fmtScheme/.test(p.path)));}finally{s.dispose();}
  for(const scheme of['<a:fmtScheme name="Empty"/>',formatScheme({fills:['<a:gradFill/>'],lines:[]})]){
    const p=await open(matrixOverride(matrixFixture(),scheme));try{assert.equal(cell(p).style.fill.type,"none");assert.equal(cell(p).style.stroke.paint.type,"none");assert.ok(properties(p).filter(p=>p.property.startsWith("cell.")).every(p=>p.status==="unsupported"));}finally{p.dispose();}
  }
  const noFormat=await open(matrixOverride(matrixFixture(),"",'<a:clrScheme name="Only colors"><a:accent1><a:srgbClr val="AABBCC"/></a:accent1><a:accent2><a:srgbClr val="DDEEFF"/></a:accent2></a:clrScheme>'));try{assert.equal(cell(noFormat).style.fill.color,"AABBCC");assert.equal(cell(noFormat).style.stroke.paint.color,"DDEEFF");}finally{noFormat.dispose();}
});

test("unsupported positional entries never shift indices; MCE is evaluated before selecting the slot",async()=>{
  const scheme=formatScheme({fills:['<x:future xmlns:x="urn:future"/>',matrixFill("445566","","srgbClr")]});
  const first=await open(matrixFixture({scheme}));try{assert.equal(cell(first).style.fill.type,"none");}finally{first.dispose();}
  const second=await open(matrixFixture({scheme,fill:fillReference(2)}));try{assert.equal(cell(second).style.fill.color,"445566");}finally{second.dispose();}
  const f=matrixFixture();f.parts['deck/theme.xml']=f.parts['deck/theme.xml'].replace('<a:fmtScheme ', '<a:fmtScheme xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future" ').replace('<a:fillStyleLst>',`<a:fillStyleLst><mc:AlternateContent><mc:Choice Requires="x"><a:gradFill/></mc:Choice><mc:Fallback>${matrixFill("ABCDEF","","srgbClr")}</mc:Fallback></mc:AlternateContent>`);
  const m=await open(f);try{assert.equal(cell(m).style.fill.color,"ABCDEF");assert.ok(m.report.inspection.markupCompatibility.some(p=>p.part==="/deck/theme.xml"));}finally{m.dispose();}
});

test("malformed indices/duplicates reject; missing/out-of-range and unsupported line semantics stay explicit",async()=>{
  for(const idx of["-1","1.0","4294967296","NaN"]){await assert.rejects(open(matrixFixture({fill:fillReference(idx)})),{code:"import-invalid-package"});}
  const absent=matrixFixture({fill:'<a:fillRef><a:schemeClr val="accent1"/></a:fillRef>'});await assert.rejects(open(absent),{code:"import-invalid-package"});
  const duplicate=matrixFixture({scheme:formatScheme().replace('</a:fmtScheme>','<a:fillStyleLst/></a:fmtScheme>')});await assert.rejects(open(duplicate),{code:"import-invalid-package"});
  for(const variant of[{fill:fillReference(999)},{scheme:formatScheme({lines:[matrixLine().replace('cmpd="sng"','cmpd="dbl"')]})},{scheme:formatScheme({lines:[matrixLine().replace('<a:miter lim="800000"/>','<a:round/>')]})}]){const s=await open(matrixFixture(variant));try{assert.ok(properties(s).some(p=>p.status==="unsupported"));if(variant.scheme?.includes('cmpd="dbl"'))assert.ok(s.report.findings.some(f=>f.source.part==="/deck/theme.xml"&&f.source.path?.endsWith("/@cmpd")),"matrix diagnostics retain their actual theme source location");}finally{s.dispose();}}
});

test("three scalar source saves retain theme matrices and references, while style mutations remain blocked with retry",async()=>{
  let bytes=encodeStructure(matrixFixture());
  for(let round=0;round<3;round++){
    const s=await importPptx(bytes,{projection:"content",editSaveProfile:"text-runs-v3"});try{
      const slide=s.document.slides[0],old=slide.elements[0],changed=structuredClone(old);changed.rows[0].cells[0].content[0].runs[0].text=`Matrix ${round}`;slide.replaceElement(old.id,changed);const ready=await s.prepareSave({mode:"edit"});assert.equal(ready.status,"ready");const next=await s.save(ready.plan),a=readGeneratedZip(bytes),b=readGeneratedZip(next.bytes);for(const[name,data]of a)if(name!=="deck/z.xml")assert.deepEqual(b.get(name),data);
      const bad=structuredClone(changed);bad.rows[0].cells[0].style.stroke.width=20;slide.replaceElement(old.id,bad);assert.equal((await s.prepareSave({mode:"edit"})).status,"blocked");slide.replaceElement(old.id,changed);assert.equal((await s.prepareSave({mode:"edit"})).status,"ready");bytes=next.bytes;
    }finally{s.dispose();}
  }
});

test("matrix traversal and retained dependencies have exact cumulative limits and retry independently",async()=>{
  const bytes=encodeStructure(matrixFixture()),baseline=await importPptx(bytes,{projection:"content"}),expected=normalizePresentation(baseline.document);baseline.dispose();
  for(const key of["maxTableStyleProperties","maxTableStyleDependencies","maxTableStyleWork"]){
    let lo=1,hi=2000000;while(lo<hi){const mid=Math.floor((lo+hi)/2);try{const s=await importPptx(bytes,{projection:"content",sessionLimits:{[key]:mid}});s.dispose();hi=mid;}catch(e){assert.equal(e.code,"import-limit-exceeded");lo=mid+1;}}
    assert.ok(lo>1);await assert.rejects(importPptx(bytes,{projection:"content",sessionLimits:{[key]:lo-1}}),{code:"import-limit-exceeded"});const s=await importPptx(bytes,{projection:"content",sessionLimits:{[key]:lo}});assert.deepEqual(normalizePresentation(s.document),expected);s.dispose();
  }
  const controller=new AbortController(),pending=importPptx(bytes,{projection:"content",signal:controller.signal});queueMicrotask(()=>controller.abort());await assert.rejects(pending,{code:"operation-aborted"});const retry=await open(matrixFixture());retry.dispose();
});

test("one shared table style selects each master's own format matrix without cross-scope values or audit state",async()=>{
  const f=matrixFixture();f.parts['deck/a.xml']=f.parts['deck/z.xml'];f.parts['deck/_rels/a.xml.rels']=f.parts['deck/_rels/a.xml.rels'].replace('layout.xml','layout2.xml');f.parts['deck/layout2.xml']=f.parts['deck/layout.xml'];f.parts['deck/_rels/layout2.xml.rels']=f.parts['deck/_rels/layout.xml.rels'].replace('master.xml','master2.xml');f.parts['deck/master2.xml']=f.parts['deck/master.xml'];f.parts['deck/_rels/master2.xml.rels']=f.parts['deck/_rels/master.xml.rels'].replace('layout.xml','layout2.xml').replace('theme.xml','theme2.xml');
  f.parts['deck/theme2.xml']=f.parts['deck/theme.xml'].replace(/<a:fmtScheme[\s\S]*?<\/a:fmtScheme>/,formatScheme({fills:[matrixFill("778899","","srgbClr")],lines:[matrixLine(matrixFill("334455","","srgbClr"),50800,"dash")]}));
  f.parts['deck/main.data']=f.parts['deck/main.data'].replace('</p:sldMasterIdLst>','<p:sldMasterId id="2147483650" r:id="m2"/></p:sldMasterIdLst>');f.parts['deck/_rels/main.data.rels']=f.parts['deck/_rels/main.data.rels'].replace('</Relationships>',`<Relationship Id="m2" Type="${f.R}/slideMaster" Target="master2.xml"/></Relationships>`);Object.assign(f.types,{'deck/layout2.xml':'presentationml.slideLayout','deck/master2.xml':'presentationml.slideMaster','deck/theme2.xml':'theme'});
  const s=await open(f);try{const p=normalizePresentation(s.document),a=p.slides[0].elements[0].rows[0].cells[0],b=p.slides[1].elements[0].rows[0].cells[0];assert.equal(a.style.fill.color,"3579BD");assert.equal(b.style.fill.color,"778899");assert.equal(b.style.stroke.width,4);assert.equal(b.style.stroke.dash,"dash");assert.equal(b.style.stroke.paint.color,"334455");assert.ok(properties(s).some(p=>p.target.part==="/deck/a.xml"&&p.resolvedFrom?.some(d=>d.part==="/deck/theme2.xml"&&/fmtScheme/.test(d.path))));}finally{s.dispose();}
  f.parts['deck/theme.xml']=f.parts['deck/theme.xml'].replace(/<a:fillStyleLst>[\s\S]*?<\/a:fillStyleLst>/,'<a:fillStyleLst><a:gradFill/></a:fillStyleLst>');const second=await open(f);try{const p=normalizePresentation(second.document);assert.equal(p.slides[0].elements[0].rows[0].cells[0].style.fill.type,"none");assert.equal(p.slides[1].elements[0].rows[0].cells[0].style.fill.color,"778899");assert.equal(properties(second).filter(p=>p.target.part==="/deck/a.xml"&&p.property==="cell.fill"&&p.status==="unsupported").length,0);}finally{second.dispose();}
});
