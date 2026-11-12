import test from "node:test";
import assert from "node:assert/strict";
import { inspectPptxMasterPictures, DEFAULT_MASTER_PICTURE_LIMITS } from "../dist/index.js";
import { structureFixture } from "./structure-fixture.mjs";
import { pack, contentTypes, relationships } from "./opc-fixture.mjs";

function fixture(strict=false, variants=false) {
  const f=structureFixture(strict);const pic=(id,embed,extra="")=>`<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="P${id}"/><p:cNvPicPr/><p:nvPr>${extra}</p:nvPr></p:nvPicPr><p:blipFill><a:blip r:embed="${embed}"/><a:srcRect l="12345" b="20"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm rot="60000"><a:off x="12700" y="25400"/><a:ext cx="38100" cy="50800"/></a:xfrm></p:spPr></p:pic>`;
  f.parts["deck/master.xml"]=`<p:sldMaster xmlns:p="${f.P}" xmlns:a="${f.A}" xmlns:r="${f.R}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/><p:sp/><p:grpSp><p:nvGrpSpPr/><p:grpSpPr/>${pic("2","image",'<p:ph type="pic" idx="4"/>')}</p:grpSp><mc:AlternateContent><mc:Choice Requires="p">${pic("3","missing")}</mc:Choice><mc:Fallback>${pic("4","image")}</mc:Fallback></mc:AlternateContent>${variants ? pic("5","wrong")+pic("6","external").replace("r:embed=","r:link=")+(variants === "missing" ? pic("7","missing-target") : "") : ""}</p:spTree></p:cSld><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="layout"/></p:sldLayoutIdLst></p:sldMaster>`;
  if (variants === "unlisted") f.parts["deck/master.xml"] = f.parts["deck/master.xml"].replace('<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="layout"/></p:sldLayoutIdLst>', '<p:sldLayoutIdLst/>');
  if (variants === "visibility") {
    f.parts["deck/a.xml"] = f.parts["deck/a.xml"].replace('<p:sld ', '<p:sld showMasterSp="false" ');
    f.parts["deck/layout.xml"] = f.parts["deck/layout.xml"].replace('<p:sldLayout ', '<p:sldLayout showMasterSp="1" ');
  }
  f.parts["media/image.png"]=Buffer.from([137,80,78,71,13,10,26,10]);
  f.parts["deck/_rels/master.xml.rels"]=relationships([f.rel("layout","slideLayout","layout.xml"),f.rel("theme","theme","theme.xml"),f.rel("image","image","../media/image.png"), ...(variants ? [f.rel("wrong","slide","a.xml"),{...f.rel("external","image","https://example.invalid/image.png"),TargetMode:"External"},...(variants === "missing" ? [f.rel("missing-target","image","../media/missing.png")] : [])] : [])]);
  return pack(f.parts,contentTypes(undefined,[...Object.entries(f.types).map(([PartName,type])=>({PartName:`/${PartName}`,ContentType:`application/vnd.openxmlformats-officedocument.${type}+xml`})),{PartName:"/media/image.png",ContentType:"image/png"}]));
}
test("master picture inventory binds effective nested objects and raw identities in both dialects",async()=>{
 for(const strict of [false,true]){
  const data=fixture(strict),result=await inspectPptxMasterPictures(data);assert.equal(result.dialect,strict?"strict":"transitional");
  assert.equal(result.coverage,"effective-listed-masters-only");assert.equal(result.mapping,"not-performed");assert.equal(result.masters.length,1);
  const m=result.masters[0];assert.deepEqual(m.slideIds,[300,256]);assert.deepEqual(m.layoutParts,["/deck/layout.xml"]);assert.equal(m.otherDrawingCount,1);
  assert.deepEqual(m.pictures.map(p=>p.sourceObjectId),["2","3"]);const [first,second]=m.pictures;
  assert.equal(first.source.part,"/deck/master.xml");assert.ok(first.source.elementIndex>0);assert.equal(first.groupElementIndices.length,1);
  assert.deepEqual(first.placeholder,{type:"pic",index:"4"});assert.deepEqual(first.embed,{id:"image",status:"internal-image",targetPart:"/media/image.png"});
  assert.deepEqual(first.crop,{left:"12345",bottom:"20"});assert.deepEqual(first.transform,{x:"12700",y:"25400",cx:"38100",cy:"50800",rotation:"60000"});
  assert.deepEqual(second.embed,{id:"missing",status:"missing"});assert.notEqual(first.source.elementIndex,second.source.elementIndex);
  assert.ok(Object.isFrozen(result.masters[0].pictures[0].groupElementIndices));
 }
});
test("master picture inventory enforces independent work, record, string and input limits",async()=>{
 const bytes=fixture();const full=await inspectPptxMasterPictures(bytes);
 for(const [key,value] of [["maxWorkUnits",full.workUnits-1],["maxRecords",full.recordCount-1],["maxTextCharacters",full.textCharacters-1]]) await assert.rejects(inspectPptxMasterPictures(bytes,{masterPictureLimits:{[key]:value}}),{code:"package-limit-exceeded"});
 for(const value of [0,-1,Infinity,1.5]) await assert.rejects(inspectPptxMasterPictures(bytes,{masterPictureLimits:{maxRecords:value}}),RangeError);
 await assert.rejects(inspectPptxMasterPictures(bytes,{masterPictureLimits:{unknown:1}}),TypeError);
 await assert.rejects(inspectPptxMasterPictures(bytes,{masterPictureLimits:{get maxRecords(){throw Error("getter invoked")}}}),TypeError);
 const controller=new AbortController();controller.abort();await assert.rejects(inspectPptxMasterPictures(bytes,{signal:controller.signal}),{code:"operation-aborted"});
 assert.ok(Object.isFrozen(DEFAULT_MASTER_PICTURE_LIMITS));
});


test("master picture relationships distinguish wrong, external and missing targets without fetching",async()=>{
 const result=await inspectPptxMasterPictures(fixture(false,true));const items=result.masters[0].pictures;
 assert.deepEqual(items.map(p=>p.sourceObjectId),["2","3","5","6"]);
 assert.deepEqual(items[2].embed,{id:"wrong",status:"wrong-type"});
 assert.deepEqual(items[3].link,{id:"external",status:"external-image"});
 await assert.rejects(inspectPptxMasterPictures(fixture(false,"missing")),{code:"package-invalid"});
});

test("input snapshot and offset views are detached before asynchronous inspection",async()=>{
 const data=fixture(),backing=Buffer.concat([Buffer.from([1,2]),data,Buffer.from([3])]),view=backing.subarray(2,2+data.length);
 const pending=inspectPptxMasterPictures(view);view.fill(0);const result=await pending;
 assert.deepEqual(result.masters[0].pictures.map(p=>p.sourceObjectId),["2","3"]);
 if(typeof SharedArrayBuffer!=="undefined") await assert.rejects(inspectPptxMasterPictures(new Uint8Array(new SharedArrayBuffer(8))),TypeError);
});


test("unlisted layout references are not claimed as verified master membership",async()=>{
 const result=await inspectPptxMasterPictures(fixture(false,"unlisted"));
 assert.deepEqual(result.masters[0].layoutParts,[]);assert.deepEqual(result.masters[0].slideIds,[]);
 assert.deepEqual(result.masters[0].slideUses,[]);
});

test("master uses retain explicit slide and layout switches without inferring visibility",async()=>{
 for (const strict of [false,true]) {
  const result=await inspectPptxMasterPictures(fixture(strict,"visibility"));
  assert.deepEqual(result.masters[0].slideUses,[
   {slideId:300,slidePart:"/deck/z.xml",layoutPart:"/deck/layout.xml",layoutShowMasterSp:"1"},
   {slideId:256,slidePart:"/deck/a.xml",layoutPart:"/deck/layout.xml",slideShowMasterSp:"false",layoutShowMasterSp:"1"},
  ]);
  assert.ok(Object.isFrozen(result.masters[0].slideUses[0]));
 }
});
