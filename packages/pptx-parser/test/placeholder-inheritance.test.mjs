import test from "node:test";
import assert from "node:assert/strict";
import { importPptx, DEFAULT_SESSION_LIMITS } from "../dist/index.js";
import { normalizePresentation } from "../../core/dist/index.js";
import { relationships } from "./opc-fixture.mjs";
import { encodeStructure } from "./structure-fixture.mjs";
import { shape, body } from "./content-fixture.mjs";
import { inheritedShape, placeholderTransform as xfrm, placeholderFixture, placeholderBytes } from "./placeholder-fixture.mjs";

const open = f => importPptx(encodeStructure(f), { projection: "content" });
const entry = (session, part = "/deck/z.xml", sourceId = "2") => session.report.placeholderInheritance.placeholders.find(p => p.source.part === part && p.source.sourceObjectId === sourceId);
const elements = s => normalizePresentation(s.document).slides[0].elements;
const errorCode = code => e => e.code === code;

test("placeholder dimensions inherit through layout and master, with exact per-property source provenance", async () => {
  for (const strict of [false, true]) {
    const s = await importPptx(placeholderBytes(strict), { projection: "content" });
    const ph = entry(s), [element] = elements(s);
    assert.equal(ph.parent.rule, "index"); assert.equal(ph.parent.status, "matched"); assert.equal(ph.parent.key, 0);
    assert.equal(ph.explicitIndex, false); assert.equal(ph.explicitType, true);
    assert.equal(ph.geometryStatus, "resolved");
    assert.equal(ph.dimensions.x.valueEmu, 127000); assert.equal(ph.dimensions.x.level, "layout");
    assert.equal(ph.dimensions.width.valueEmu, 2540000); assert.equal(ph.dimensions.width.level, "master");
    assert.match(ph.dimensions.width.source.path, /a:ext\[1\]\/@cx$/);
    assert.equal(ph.dimensions.width.source.sourceObjectId, "2");
    assert.deepEqual(element.box, { x: 10, y: 20, width: 200, height: 100 });
    assert.equal(element.text.plainText, " A & B ");
    assert.equal(s.report.placeholderInheritance.placeholders.length, 4, "shared layout/master declarations appear once");
    assert.equal(entry(s, "/deck/layout.xml").parent.rule, "type");
    assert.equal(entry(s, "/deck/layout.xml").parent.source.part, "/deck/master.xml");
    const preset = s.report.findings.find(f => f.feature === "drawing.placeholder-preset");
    assert.equal(preset.source.part, "/deck/master.xml");
    assert.match(preset.source.path, /a:prstGeom\[1\]$/);
    assert.equal(preset.source.sourceObjectId, "2");
    assert.equal(s.document.layouts.length, 0, "source masters are not fabricated Core layouts");
    assert.equal(JSON.stringify(normalizePresentation(s.document)).includes("MASTER PROMPT"), false);
    assert.equal(s.report.identities.filter(i => i.kind === "element").length, 2);
    assert.ok(Object.isFrozen(ph.dimensions.width.source)); assert.ok(Object.isFrozen(ph.parent));
    s.dispose();
  }
});

test("slide matches numeric index instead of type, drawing ID, name or array order", async () => {
  const f = placeholderFixture();
  f.set("deck/z.xml", "sld", inheritedShape({ ph: 'type="body" idx="+0007"', preset: "rect" }));
  f.set("deck/layout.xml", "sldLayout",
    inheritedShape({ id: "9", ph: 'type="body" idx="0"', xfrm: xfrm({ x: 0 }) }) +
    inheritedShape({ id: "4", ph: 'type="title" idx="7"', xfrm: xfrm({ x: 1270000 }), preset: "rect" }));
  const s = await open(f), ph = entry(s);
  assert.equal(ph.index, 7); assert.equal(ph.parent.source.sourceObjectId, "4");
  assert.equal(elements(s)[0].box.x, 100); s.dispose();
});

test("layout maps title/center-title and content placeholder families to the owning master type", async () => {
  const families = { title: "title", ctrTitle: "title", body: "body", subTitle: "body", obj: "body", chart: "body", tbl: "body", dgm: "body", media: "body", clipArt: "body", pic: "body", dt: "dt", ftr: "ftr", sldNum: "sldNum" };
  for (const [type, masterType] of Object.entries(families)) {
    const f = placeholderFixture();
    f.set("deck/z.xml", "sld", inheritedShape({ ph: 'idx="5"', preset: "rect" }));
    f.set("deck/layout.xml", "sldLayout", inheritedShape({ ph: `type="${type}" idx="5"` }));
    f.set("deck/master.xml", "sldMaster", inheritedShape({ ph: `type="${masterType}" idx="99"`, xfrm: xfrm({ x: 381000 }), preset: "rect" }),
      '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="layout"/></p:sldLayoutIdLst>');
    const s = await open(f);
    assert.equal(entry(s, "/deck/layout.xml").parent.key, masterType, type);
    assert.equal(entry(s, "/deck/layout.xml").parent.status, "matched", type);
    assert.equal(elements(s)[0].box.x, 30, type);
    s.dispose();
  }
});

test("omitted type/index use obj/0 while explicit false/zero geometry is never replaced by ancestor values", async () => {
  const f = placeholderFixture();
  f.set("deck/z.xml", "sld", inheritedShape({ ph: "", xfrm: xfrm({ x: 0, y: 0, ext: false, attrs: 'rot="0" flipH="0" flipV="false"' }), preset: "rect" }));
  f.set("deck/layout.xml", "sldLayout", inheritedShape({ ph: "", xfrm: xfrm({ width: 0, height: 0 }) }));
  const s = await open(f), ph = entry(s);
  assert.equal(ph.type, "obj"); assert.equal(ph.index, 0);
  assert.equal(ph.explicitType, false); assert.equal(ph.explicitIndex, false);
  assert.deepEqual(elements(s)[0].box, { x: 0, y: 0, width: 0, height: 0 });
  assert.deepEqual(elements(s)[0].transform, { rotation: 0, flipH: false, flipV: false });
  assert.equal(ph.dimensions.x.level, "slide"); assert.equal(ph.dimensions.width.level, "layout");
  s.dispose();
});

test("duplicate layout indices are ambiguous, including lexical aliases, with no first-candidate guess", async () => {
  const f = placeholderFixture();
  f.set("deck/layout.xml", "sldLayout", inheritedShape({ ph: 'type="title" idx="0"', xfrm: xfrm(), preset: "rect" }) + inheritedShape({ id: "3", ph: 'type="body" idx="+00"', xfrm: xfrm({ x: 999 }), preset: "rect" }));
  const s = await open(f), ph = entry(s);
  assert.equal(ph.parent.status, "ambiguous"); assert.equal(ph.parent.candidateCount, 2);
  assert.equal(ph.parent.source, undefined); assert.equal(ph.geometryStatus, "incomplete");
  assert.deepEqual(elements(s), []); assert.equal(s.report.identities.some(i => i.kind === "element"), false);
  assert.ok(s.report.findings.filter(f => f.source.part === "/deck/z.xml").every(f => f.elementId === undefined), "omitted drawing findings do not manufacture Core identities");
  assert.ok(s.report.findings.some(f => f.feature === "drawing.placeholder-match" && f.tier === "fallback"));
  assert.equal(s.report.placeholderInheritance.placeholders.length, 5); s.dispose();
});

test("omitted slide and layout indexes remain ambiguous even when placeholder roles differ", async () => {
  const f = placeholderFixture();
  f.set("deck/z.xml", "sld", inheritedShape({ id: "2", ph: 'type="title"', xfrm: xfrm({ x: 127000 }), preset: "rect" }) + inheritedShape({ id: "3", ph: "", xfrm: xfrm({ x: 254000 }), preset: "rect" }));
  f.set("deck/layout.xml", "sldLayout", inheritedShape({ id: "2", ph: 'type="title"', xfrm: xfrm({ x: 381000 }), preset: "rect" }) + inheritedShape({ id: "3", ph: 'type="body"', xfrm: xfrm({ x: 508000 }), preset: "rect" }));
  const bytes = encodeStructure(f), session = await importPptx(bytes, { projection: "content" });
  try {
    for (const id of ["2", "3"]) {
      const ph = entry(session, "/deck/z.xml", id);
      assert.equal(ph.index, 0); assert.equal(ph.explicitIndex, false);
      assert.equal(ph.parent.status, "ambiguous"); assert.equal(ph.parent.candidateCount, 2);
      assert.equal(ph.parent.source, undefined); assert.equal(ph.geometryStatus, "resolved");
    }
    assert.deepEqual(elements(session).map(item => item.box.x), [10, 20]);
    assert.equal(session.report.findings.filter(finding => finding.feature === "drawing.placeholder-match" && finding.tier === "fallback" && finding.source.part === "/deck/z.xml").length, 2);
    const plan = await session.prepareSave({ mode: "no-op" }); assert.equal(plan.status, "ready");
    assert.deepEqual((await session.save(plan.plan)).bytes, new Uint8Array(bytes));
  } finally { session.dispose(); }
});

test("duplicate master types block inherited values; complete local geometry still projects without parent guessing", async () => {
  const f = placeholderFixture();
  f.set("deck/master.xml", "sldMaster", inheritedShape({ xfrm: xfrm(), preset: "rect" }) + inheritedShape({ id: "3", xfrm: xfrm(), preset: "rect" }),
    '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="layout"/></p:sldLayoutIdLst>');
  f.set("deck/z.xml", "sld", inheritedShape({ xfrm: xfrm({ x: -12700 }), preset: "rect", props: '<a:ln><a:noFill/></a:ln>' }));
  const s = await open(f);
  assert.equal(entry(s, "/deck/layout.xml").parent.status, "ambiguous");
  assert.equal(entry(s).geometryStatus, "resolved");
  assert.equal(elements(s)[0].box.x, -1); s.dispose();
});

test("missing relationships, missing matches and unsupported header types are separately reported", async () => {
  for (const mode of ["no-layout", "missing-index", "header"]) {
    const f = placeholderFixture();
    if (mode === "no-layout") f.parts["deck/_rels/z.xml.rels"] = relationships([]);
    if (mode === "missing-index") f.set("deck/z.xml", "sld", inheritedShape({ ph: 'idx="123"' }));
    if (mode === "header") f.set("deck/layout.xml", "sldLayout", inheritedShape({ ph: 'type="hdr"' }));
    const s = await open(f);
    const match = entry(s, mode === "header" ? "/deck/layout.xml" : "/deck/z.xml").parent;
    assert.equal(match.status, mode === "header" ? "unsupported" : "missing", mode);
    assert.deepEqual(elements(s), []); s.dispose();
  }
});

test("ordinary shapes are not matched by coincident names or IDs, and unknown/custom geometry blocks inherited presets", async () => {
  for (const content of [
    shape({ xfrm: "" }),
    inheritedShape({ props: "<a:custGeom/>" }),
    inheritedShape({ preset: "cloud" }),
  ]) {
    const f = placeholderFixture(); f.set("deck/z.xml", "sld", content);
    const s = await open(f);
    assert.deepEqual(elements(s), []);
    if (content.startsWith("<p:sp>") && !content.includes("<p:ph")) assert.equal(entry(s), undefined);
    s.dispose();
  }
});

test("multi-master matching is scoped through source relationships and never crosses to an unrelated theme/master", async () => {
  const f = placeholderFixture();
  f.parts["deck/main.data"] = f.parts["deck/main.data"].replace("</p:sldMasterIdLst>", '<p:sldMasterId id="2147483650" r:id="m2"/></p:sldMasterIdLst>');
  f.parts["deck/_rels/main.data.rels"] = relationships([f.rel("a", "slide", "a.xml"), f.rel("m", "slideMaster", "master.xml"), f.rel("z", "slide", "z.xml"), f.rel("m2", "slideMaster", "master2.xml")]);
  f.parts["deck/layout2.xml"] = f.parts["deck/layout.xml"];
  f.parts["deck/master2.xml"] = f.parts["deck/master.xml"].replace('cx="2540000"', 'cx="3810000"');
  f.parts["deck/theme2.xml"] = f.parts["deck/theme.xml"];
  f.parts["deck/_rels/layout2.xml.rels"] = relationships([f.rel("master", "slideMaster", "master2.xml")]);
  f.parts["deck/_rels/master2.xml.rels"] = relationships([f.rel("layout", "slideLayout", "layout2.xml"), f.rel("theme", "theme", "theme2.xml")]);
  f.parts["deck/_rels/a.xml.rels"] = relationships([f.rel("layout", "slideLayout", "layout2.xml")]);
  for (const [target, source] of [["deck/layout2.xml", "deck/layout.xml"], ["deck/master2.xml", "deck/master.xml"], ["deck/theme2.xml", "deck/theme.xml"]]) f.types[target] = f.types[source];
  const s = await open(f), ir = normalizePresentation(s.document);
  assert.equal(ir.slides[0].elements[0].box.width, 200); assert.equal(ir.slides[1].elements[0].box.width, 300);
  assert.equal(entry(s, "/deck/a.xml").dimensions.width.source.part, "/deck/master2.xml");
  assert.equal(s.report.placeholderInheritance.placeholders.length, 6);
  s.document.slides[0].removeElement(s.document.slides[0].elements[0].id);
  assert.equal(normalizePresentation(s.document).slides[1].elements[0].box.width, 300);
  s.dispose();
});

function group(content, id = "3") {
  return `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${id}" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1270000" cy="1270000"/><a:chOff x="0" y="0"/><a:chExt cx="1270000" cy="1270000"/></a:xfrm></p:grpSpPr>${content}</p:grpSp>`;
}
test("grouped source or candidate placeholders retain explicit coordinate limitations", async () => {
  for (const parentGrouped of [false, true]) {
    const f = placeholderFixture();
    if (parentGrouped) f.set("deck/layout.xml", "sldLayout", group(inheritedShape({ xfrm: xfrm(), preset: "rect" })));
    else f.set("deck/z.xml", "sld", group(inheritedShape()));
    const s = await open(f);
    assert.equal(entry(s).parent.status, "unsupported");
    assert.equal(parentGrouped ? elements(s).length : elements(s)[0].children.length, 0);
    assert.match(entry(s).parent.reason, /Grouped/); s.dispose();
  }
});

test("ambiguous parent classification does not depend on whether the first candidate is grouped", async () => {
  for (const reverse of [false, true]) {
    const f = placeholderFixture();
    const candidates = [group(inheritedShape({ xfrm: xfrm(), preset: "rect" })), inheritedShape({ id: "4", xfrm: xfrm(), preset: "rect" })];
    f.set("deck/layout.xml", "sldLayout", (reverse ? candidates.reverse() : candidates).join(""));
    const s = await open(f);
    assert.equal(entry(s).parent.status, "ambiguous");
    assert.equal(entry(s).parent.candidateCount, 2);
    assert.equal(entry(s).parent.source, undefined);
    assert.deepEqual(elements(s), []); s.dispose();
  }
});

test("ancestor rotation/reflection is recorded as unsupported instead of silently projected as a neutral transform", async () => {
  for (const attrs of ['rot="5400000"', 'flipH="true"', 'flipV="1"']) {
    const f = placeholderFixture();
    f.set("deck/layout.xml", "sldLayout", inheritedShape({ xfrm: xfrm({ attrs }) }));
    const s = await open(f);
    assert.equal(entry(s).geometryStatus, "unsupported");
    assert.match(entry(s).reason, /rotation\/reflection/); assert.deepEqual(elements(s), []); s.dispose();
  }
  const f = placeholderFixture();
  f.set("deck/z.xml", "sld", inheritedShape({ xfrm: xfrm({ attrs: 'rot="5400000" flipH="true"', ext: false }), preset: "rect" }));
  const s = await open(f);
  assert.deepEqual(elements(s)[0].transform, { rotation: 90, flipH: true, flipV: false }); s.dispose();
  const cleared = placeholderFixture();
  cleared.set("deck/layout.xml", "sldLayout", inheritedShape({ xfrm: xfrm({ attrs: 'rot="5400000" flipH="1" flipV="true"' }) }));
  cleared.set("deck/z.xml", "sld", inheritedShape({ xfrm: xfrm({ x: 0, y: 0, attrs: 'rot="0" flipH="false" flipV="0"' }), preset: "rect" }));
  const local = await open(cleared);
  assert.equal(entry(local).geometryStatus, "resolved");
  assert.deepEqual(elements(local)[0].transform, { rotation: 0, flipH: false, flipV: false });
  assert.equal(elements(local)[0].box.x, 0); local.dispose();
});

test("Strict aliases and effective MCE branches match while opaque extension placeholders remain invisible", async () => {
  const f = placeholderFixture(true), original = inheritedShape({ xfrm: xfrm({ ext: false }) });
  f.parts["deck/layout.xml"] = f.parts["deck/layout.xml"]
    .replace(original, `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:u="urn:unknown"><mc:Choice Requires="u">${inheritedShape({ ph: 'idx="6"' })}</mc:Choice><mc:Fallback>${original}</mc:Fallback></mc:AlternateContent>`)
    .replace("</p:cSld>", `<p:extLst><p:ext uri="urn:opaque">${inheritedShape({ id: "2" })}</p:ext></p:extLst></p:cSld>`)
    .replaceAll("xmlns:p=", "xmlns:s=").replaceAll("<p:", "<s:").replaceAll("</p:", "</s:");
  const s = await open(f);
  assert.equal(elements(s).length, 1); assert.equal(entry(s, "/deck/layout.xml").index, 0);
  assert.match(entry(s).dimensions.x.source.path, /s:sp\[1\]/);
  assert.equal(s.report.placeholderInheritance.placeholders.length, 4);
  assert.ok(s.report.inspection.markupCompatibility.some(e => e.part === "/deck/layout.xml")); s.dispose();
});

test("malformed indices/types, duplicate declarations and malformed inherited geometry fail before publishing a session", async () => {
  for (const ph of ['idx="-1"', 'idx="4294967296"', 'idx=""', 'idx="1e2"', 'type="unknown"', 'type=""']) {
    const f = placeholderFixture(); f.set("deck/z.xml", "sld", inheritedShape({ ph }));
    await assert.rejects(open(f), errorCode("import-invalid-package"));
  }
  for (const content of [
    inheritedShape().replace("<p:ph", "<p:ph/><p:ph"),
    inheritedShape({ xfrm: xfrm().replace('cx="1270000"', 'cx="-1"') }),
    inheritedShape({ xfrm: xfrm({ attrs: 'flipH="maybe"' }) }),
    inheritedShape({ xfrm: xfrm().replace('x="127000"', 'x="not-a-number"') }),
    inheritedShape() + inheritedShape(),
  ]) {
    const f = placeholderFixture(); f.set("deck/master.xml", "sldMaster", content, '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="layout"/></p:sldLayoutIdLst>');
    await assert.rejects(open(f), errorCode("import-invalid-package"));
  }
});

test("provenance budget counts shared declarations once, is snapshotted, and cancellation/concurrent sessions remain isolated", async () => {
  assert.equal(DEFAULT_SESSION_LIMITS.maxPlaceholderEntries, 100000);
  const input = placeholderBytes();
  await assert.rejects(importPptx(input, { projection: "content", sessionLimits: { maxPlaceholderEntries: 3 } }), errorCode("import-limit-exceeded"));
  const options = { projection: "content", sessionLimits: { maxPlaceholderEntries: 4 } };
  const pending = importPptx(input, options); options.sessionLimits.maxPlaceholderEntries = 1;
  const a = await pending, b = await importPptx(input, { projection: "content", sessionLimits: { maxPlaceholderEntries: 4 } });
  assert.deepEqual(a.report.placeholderInheritance, b.report.placeholderInheritance);
  a.dispose(); assert.equal(entry(b).dimensions.width.valueEmu, 2540000); b.dispose();
  for (const value of [0, -1, 1.5, NaN, Infinity]) await assert.rejects(importPptx(input, { projection: "content", sessionLimits: { maxPlaceholderEntries: value } }), RangeError);
  const controller = new AbortController(), aborted = importPptx(input, { projection: "content", signal: controller.signal });
  controller.abort(); await assert.rejects(aborted, errorCode("operation-aborted"));
});

test("exact-byte no-op, stale plans and structure mode retain separate guarantees", async () => {
  const input = placeholderBytes(), s = await importPptx(input, { projection: "content" });
  const prepared = await s.prepareSave({ mode: "no-op" });
  assert.deepEqual((await s.save(prepared.plan)).bytes, new Uint8Array(input));
  assert.equal((await s.prepareSave({ mode: "edit" })).status, "blocked");
  assert.equal((await s.prepareSave({ mode: "rebuild" })).status, "blocked");
  const slide = s.document.slides[0], previous = slide.removeElement(slide.elements[0].id);
  await assert.rejects(s.save(prepared.plan), errorCode("save-snapshot-stale"));
  slide.addElement(previous); assert.deepEqual((await s.save(prepared.plan)).bytes, new Uint8Array(input));
  s.dispose();
  const structure = await importPptx(input, { projection: "structure", sessionLimits: { maxPlaceholderEntries: 1 } });
  assert.equal(structure.report.placeholderInheritance, undefined);
  assert.equal(structure.document.slides[0].elements.length, 0); structure.dispose();
});

test('bounded inherited parallelogram slant projects with provenance while explicit local geometry wins', async () => {
  for (const strict of [false,true]) {
    const f=placeholderFixture(strict);
    const adjusted=inheritedShape({xfrm:xfrm(),preset:'parallelogram'}).replace('<a:avLst/>','<a:avLst><a:gd name="adj" fmla="val 40000"/></a:avLst>');
    f.set('deck/layout.xml','sldLayout',adjusted);
    f.set('deck/z.xml','sld',inheritedShape({text:body()}));
    let session=await open(f);
    try{assert.equal(session.document.slides[0].elements[0].shape,'parallelogram');assert.equal(session.document.slides[0].elements[0].slant,.4);const finding=session.report.findings.find(x=>x.feature==='drawing.placeholder-slant');assert.equal(finding.source.part,'/deck/layout.xml');assert.match(finding.source.path,/a:gd\[1\]$/);}finally{session.dispose();}
    f.set('deck/z.xml','sld',inheritedShape({preset:'parallelogram',text:body()}));
    session=await open(f);
    try{assert.equal(session.document.slides[0].elements[0].shape,'parallelogram');assert.equal(session.document.slides[0].elements[0].slant,undefined);}finally{session.dispose();}
  }
});

test('inherited slant selects nearest matched geometry and blocks unsupported or ambiguous sources', async () => {
  for(const strict of [false,true]){
    const f=placeholderFixture(strict), make=(value,id='2')=>inheritedShape({id,xfrm:xfrm(),preset:'parallelogram'}).replace('<a:avLst/>',`<a:avLst><a:gd name="adj" fmla="${value}"/></a:avLst>`);
    f.set('deck/master.xml','sldMaster',make('val 60000'),'<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="layout"/></p:sldLayoutIdLst>');
    f.set('deck/layout.xml','sldLayout',make('val 40000'));f.set('deck/z.xml','sld',inheritedShape({text:body()}));
    let session=await open(f);try{assert.equal(session.document.slides[0].elements[0].slant,.4);assert.equal(session.report.findings.find(x=>x.feature==='drawing.placeholder-slant').source.part,'/deck/layout.xml');}finally{session.dispose();}
    f.set('deck/layout.xml','sldLayout',inheritedShape({xfrm:xfrm()}));session=await open(f);try{assert.equal(session.document.slides[0].elements[0].slant,.6);assert.equal(session.report.findings.find(x=>x.feature==='drawing.placeholder-slant').source.part,'/deck/master.xml');}finally{session.dispose();}
    f.set('deck/layout.xml','sldLayout',make('*/ 1 2 3'));session=await open(f);try{assert.equal(session.document.slides[0].elements.length,0);assert.ok(session.report.findings.some(x=>x.reason.includes('Parallelogram adjustments')));}finally{session.dispose();}
    f.set('deck/layout.xml','sldLayout',make('val 40000')+make('val 50000','3'));session=await open(f);try{assert.equal(session.document.slides[0].elements.length,0);assert.ok(!session.report.findings.some(x=>x.feature==='drawing.placeholder-slant'));}finally{session.dispose();}
  }
});
