import assert from "node:assert/strict";
import test from "node:test";
import { mkdtemp, writeFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { importPptx, DEFAULT_SESSION_LIMITS } from "../dist/index.js";
import { importPptxFile } from "../dist/node.js";
import { diagramPreviewFixture, cacheTextTransform } from "./diagram-preview-fixture.mjs";
import { transform } from "./content-fixture.mjs";
import { relationships } from "./opc-fixture.mjs";

const options = { projection: "content", diagramPreview: "cache", editSaveProfile: "text-runs-v3" };
const leaf = session => session.document.slides[0].elements[0].children[0].children[0];
test("cache projection is explicit, captured at invocation and restricted to content mode", async () => {
  const f = diagramPreviewFixture(), normal = await importPptx(f.bytes(), { projection: "content" });
  assert.equal(normal.document.slides[0].elements.length, 1); assert.equal(normal.report.diagramPreviews, undefined); normal.dispose();
  for (const invalid of [{ projection: "structure", diagramPreview: "cache" }, { projection: "content", diagramPreview: "other" }]) await assert.rejects(importPptx(f.bytes(), invalid), TypeError);
  const config = { ...options }, pending = importPptx(f.bytes(), config); config.diagramPreview = undefined;
  const s = await pending; assert.equal(s.report.diagramPreviews.length, 1); s.dispose();
});

test("both dialects map cache shape/text geometry as a fallback, with no native node/edit permission", async () => {
  for (const strict of [false, true]) {
    const f = diagramPreviewFixture(strict), bytes = f.bytes(), s = await importPptx(bytes, options);
    try {
      assert.equal(s.report.diagramPreviews[0].status, "cache"); assert.equal(s.report.diagramPreviews[0].freshness, "unknown");
      assert.equal(s.document.slides[0].elements[0].type, "group"); assert.deepEqual(s.document.slides[0].elements[0].box, { x: 100, y: 50, width: 400, height: 200 });
      assert.equal(leaf(s).type, "shape"); assert.equal(leaf(s).text.content[0].runs[0].text, "Cached text");
      assert.deepEqual(leaf(s).box, { x: 10, y: 20, width: 100, height: 50 });
      assert.ok(s.report.findings.some(f => f.feature === "smartart.cache-preview" && f.operation === "preview" && f.tier === "fallback"));
      assert.ok(!s.report.textRuns.some(r => r.elementId.startsWith(s.document.slides[0].elements[0].id)));
      assert.ok(!s.report.findings.some(f => f.feature === "smartart.cache-preview" && f.tier === "editable"));
      const prep = await s.prepareSave({ mode: "no-op" }); assert.equal(prep.status, "ready"); assert.deepEqual((await s.save(prep.plan)).bytes, new Uint8Array(bytes));
    } finally { s.dispose(); }
  }
});

test("root/nested group child origins and local flips/rotation are represented once", async () => {
  const f = diagramPreviewFixture();
  const group = '<dsp:grpSp><dsp:nvGrpSpPr/><dsp:grpSpPr><a:xfrm rot="5400000" flipH="1"><a:off x="127000" y="254000"/><a:ext cx="2540000" cy="1270000"/><a:chOff x="127000" y="254000"/><a:chExt cx="1270000" cy="635000"/></a:xfrm></dsp:grpSpPr>' + f.cacheShape() + '</dsp:grpSp>';
  f.setCache(group); const s = await importPptx(f.bytes(), options);
  try { const g = leaf(s); assert.equal(g.type, "group"); assert.equal(g.transform.rotation, 90); assert.equal(g.transform.flipH, true); assert.deepEqual(g.coordinateSize, { width: 100, height: 50 }); assert.equal(g.children[0].box.x, 0); assert.equal(g.children[0].box.y, 0); } finally { s.dispose(); }
  f.setCache(f.cacheShape(), '<dsp:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="5080000" cy="2540000"/><a:chOff x="127000" y="254000"/><a:chExt cx="1270000" cy="635000"/></a:xfrm></dsp:grpSpPr>');
  const root = await importPptx(f.bytes(), options);
  try { assert.deepEqual(root.document.slides[0].elements[0].children[0].coordinateSize, { width: 100, height: 50 }); assert.equal(leaf(root).box.x, 0); assert.equal(leaf(root).box.y, 0); } finally { root.dispose(); }
});

test("missing, empty, source-only and invalid-root caches yield located placeholder reports", async () => {
  for (const [change, reason] of [
    [f => f.setData(""), /No located cache/], [f => f.setCache(""), /empty/],
    [f => f.setCache(f.cacheShape(), '<dsp:grpSpPr/>', 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'), /MCE/],
    [f => { f.parts["deck/diagrams/cache.xml"] = '<drawing xmlns="urn:unknown"/>'; }, /wrong-root/],
  ]) {
    const f = diagramPreviewFixture(); change(f); const s = await importPptx(f.bytes(), options);
    assert.equal(s.report.diagramPreviews[0].status, "placeholder"); assert.match(s.report.diagramPreviews[0].reason, reason); assert.equal(s.document.slides[0].elements[0].children[0].text.content, "SmartArt"); s.dispose();
  }
});

test("unknown geometry/independent text transforms and source actions stay explicit visual fallbacks", async () => {
  for (const props of [{ preset: "star5" }, { extra: '<dsp:txXfrm/>' }, { preset: "star5", hidden: true }]) {
    const f = diagramPreviewFixture(); f.setCache(f.cacheShape("2", props)); const s = await importPptx(f.bytes(), options);
    assert.equal(s.report.diagramPreviews[0].status, "partial-cache");
    if (props.extra) { assert.equal(leaf(s).type, "group"); assert.equal(leaf(s).children[0].type, "shape"); assert.equal(leaf(s).children[1].text.content, "SmartArt"); }
    else assert.equal(leaf(s).text.content, "SmartArt");
    if (props.hidden) assert.equal(leaf(s).hidden, true);
    assert.ok(s.report.findings.some(f => f.feature === "smartart.cache.unmapped-property")); s.dispose();
  }
});

test("both dialects keep independent text bounds and compensate shape rotation without rotating the wrapper", async () => {
  for (const strict of [false, true]) for (const [shapeAngle, textAngle, expected] of [[90, -90, 0], [345, -345, 0], [30, 15, 45], [-30, -15, 315], [12.5, -2.25, 10.25], [720, -450, 270]]) {
    const f = diagramPreviewFixture(strict);
    f.setCache(f.cacheShape("2", { xfrm: transform(`rot="${shapeAngle * 60000}"`), extra: cacheTextTransform(`rot="${textAngle * 60000}"`) }));
    const s = await importPptx(f.bytes(), options);
    try {
      const g = leaf(s); assert.equal(g.type, "group"); assert.equal(g.transform.rotation, 0);
      assert.equal(g.children[0].transform.rotation, shapeAngle); assert.equal(g.children[0].text, undefined);
      assert.deepEqual(g.children[1].box, { x: 30, y: 25, width: 80, height: 20 }); assert.equal(g.children[1].transform.rotation, expected);
      assert.equal(g.children[1].content[0].runs[0].text, "Cached text"); assert.equal(s.report.diagramPreviews[0].status, "cache");
      assert.ok(!s.report.textRuns.some(r => r.elementId === g.children[1].id));
      const slide = s.document.slides[0], copy = structuredClone(slide.elements[0]); copy.children[0].children[0].children[1].content[0].runs[0].text = "Unbound edit";
      slide.removeElement(copy.id); slide.insertElement(0, copy); assert.equal((await s.prepareSave({ mode: "edit" })).status, "blocked");
    } finally { s.dispose(); }
  }
  const imageFixture = diagramPreviewFixture(); imageFixture.setCache(imageFixture.cacheShape("2", { fill: '<a:blipFill><a:blip r:embed="cache-image"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>', extra: cacheTextTransform() }));
  const imageSession = await importPptx(imageFixture.bytes(), options);
  try { assert.equal(leaf(imageSession).children[0].children[0].type, "image"); assert.equal(leaf(imageSession).children[1].type, "text"); assert.equal(imageSession.document.assets.length, 1); } finally { imageSession.dispose(); }
});

test("independent text survives unknown visual presets and accounts for parent child origins and body rotation", async () => {
  const f = diagramPreviewFixture();
  const shape = f.cacheShape("2", { preset: "gear6", xfrm: transform('rot="5400000"'), extra: cacheTextTransform('rot="-5400000"') }).replace('<a:bodyPr ', '<a:bodyPr rot="900000" ');
  f.setCache('<dsp:grpSp><dsp:nvGrpSpPr/><dsp:grpSpPr><a:xfrm><a:off x="127000" y="254000"/><a:ext cx="2540000" cy="1270000"/><a:chOff x="127000" y="254000"/><a:chExt cx="1270000" cy="635000"/></a:xfrm></dsp:grpSpPr>' + shape + '</dsp:grpSp>');
  const s = await importPptx(f.bytes(), options);
  try {
    const g = leaf(s).children[0]; assert.equal(g.children[0].text, undefined); assert.equal(g.children[0].shape, "rect");
    assert.deepEqual(g.children[1].box, { x: 20, y: 5, width: 80, height: 20 }); assert.equal(g.children[1].transform.rotation, 15);
    assert.equal(g.children[1].content[0].runs[0].text, "Cached text"); assert.equal(s.report.diagramPreviews[0].status, "partial-cache");
  } finally { s.dispose(); }
});

test("incomplete/flipped/upright text stays unavailable and malformed mapped transforms locate the cache part", async () => {
  for (const [attrs, shapeAttrs, bodyAttrs] of [['flipH="1"', '', ''], ['', 'flipV="1"', ''], ['', '', 'upright="1"']]) {
    const f = diagramPreviewFixture();
    f.setCache(f.cacheShape("2", { xfrm: transform(shapeAttrs), extra: cacheTextTransform(attrs) }).replace('<a:bodyPr ', `<a:bodyPr ${bodyAttrs} `));
    const s = await importPptx(f.bytes(), options);
    assert.equal(leaf(s).children[1].type, "shape"); assert.equal(s.report.diagramPreviews[0].status, "partial-cache"); s.dispose();
  }
  for (const extra of [cacheTextTransform('rot="bad"'), cacheTextTransform('', { width: -1 }), cacheTextTransform() + cacheTextTransform()]) {
    const f = diagramPreviewFixture(); f.setCache(f.cacheShape("2", { extra }));
    await assert.rejects(importPptx(f.bytes(), options), e => e.code === "import-invalid-package" && e.source.part === "/deck/diagrams/cache.xml");
  }
  const opaque = diagramPreviewFixture(); opaque.setCache(opaque.cacheShape("2", { extra: cacheTextTransform('flipH="1"') }).replace('sz="1800"', 'sz="bad"'));
  const unavailable = await importPptx(opaque.bytes(), options);
  assert.equal(unavailable.report.diagramPreviews[0].status, "partial-cache"); unavailable.dispose();
});

test("cache image fills share immutable byte assets and enforce existing image/traversal budgets", async () => {
  const f = diagramPreviewFixture(); f.setCache(f.cacheShape("2", { fill: '<a:blipFill><a:blip r:embed="cache-image"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>' }));
  f.setTree(f.previewFrame + f.previewFrame.replace('id="2"', 'id="3"'));
  const s = await importPptx(f.bytes(), options);
  try { assert.equal(s.document.assets.length, 1); assert.equal(leaf(s).type, "group"); assert.equal(leaf(s).children[0].type, "image"); } finally { s.dispose(); }
  await assert.rejects(importPptx(f.bytes(), { ...options, sessionLimits: { maxImportedAssetBytes: 8 } }), { code: "import-limit-exceeded" });
  await assert.rejects(importPptx(f.bytes(), { ...options, sessionLimits: { maxDiagramPreviewWork: 1 } }), { code: "import-limit-exceeded" });
  assert.ok(DEFAULT_SESSION_LIMITS.maxDiagramPreviewWork > 1);
  const linked = diagramPreviewFixture(); linked.setCache(linked.cacheShape("2", { fill: '<a:blipFill><a:blip r:link="remote"/><a:stretch/></a:blipFill>' }));
  linked.parts["deck/diagrams/_rels/cache.xml.rels"] = relationships([{ ...linked.rel("remote", "image", "https://invalid.example/image"), TargetMode: "External" }]);
  const external = await importPptx(linked.bytes(), options);
  assert.equal(external.document.assets.length, 0); assert.equal(external.report.diagramPreviews[0].status, "partial-cache"); external.dispose();
  const repeated = diagramPreviewFixture();
  const bounded = await importPptx(repeated.bytes(), { ...options, sessionLimits: { maxDiagramPreviewWork: 250 } }); bounded.dispose();
  repeated.setTree(Array.from({ length: 20 }, (_, i) => repeated.previewFrame.replace('id="2"', `id="${i + 2}"`)).join(""));
  await assert.rejects(importPptx(repeated.bytes(), { ...options, sessionLimits: { maxDiagramPreviewWork: 250 } }), { code: "import-limit-exceeded" });
});

test("ordinary text edit saves reproject with the same preview; cache text edits are blocked", async () => {
  const f = diagramPreviewFixture(), s = await importPptx(f.bytes(), options);
  try {
    const slide = s.document.slides[0], e = structuredClone(slide.elements[1]); e.content[0].runs[0].text = "Unrelated edit";
    slide.removeElement(e.id); slide.insertElement(1, e);
    const prep = await s.prepareSave({ mode: "edit" }); assert.equal(prep.status, "ready"); await s.save(prep.plan);
    const group = structuredClone(slide.elements[0]); group.children[0].children[0].text.content[0].runs[0].text = "Cache edit";
    slide.removeElement(group.id); slide.insertElement(0, group); assert.equal((await s.prepareSave({ mode: "edit" })).status, "blocked");
  } finally { s.dispose(); }
});

test("malformed mapped cache geometry fails with cache provenance; file adapter captures the new option", async t => {
  const f = diagramPreviewFixture(); f.setCache(f.cacheShape("2", { xfrm: transform().replace('cx="1270000"', 'cx="-1"') }));
  await assert.rejects(importPptx(f.bytes(), options), e => e.code === "import-invalid-package" && e.source.part === "/deck/diagrams/cache.xml");
  const dir = await mkdtemp(path.join(tmpdir(), "pptkit-cache-preview-")); t.after(() => rm(dir, { recursive: true, force: true }));
  const file = path.join(dir, "source.pptx"); await writeFile(file, diagramPreviewFixture().bytes());
  const config = { ...options }, promise = importPptxFile(file, config); config.diagramPreview = undefined;
  const { session } = await promise; assert.equal(session.report.diagramPreviews[0].status, "cache"); session.dispose();
});

test('cache parallelogram slant maps bounded literals in both dialects and rejects formulas', async () => {
  for (const strict of [false, true]) for (const [formula, expected] of [['val 40000', .4], ['val 0', 0], ['val 100000', 1], ['*/ 1 2 3', undefined], ['val 100001', undefined]]) {
    const f=diagramPreviewFixture(strict), shape=f.cacheShape('2',{preset:'parallelogram'}).replace('<a:avLst/>', `<a:avLst><a:gd name="adj" fmla="${formula}"/></a:avLst>`);f.setCache(shape);
    const session=await importPptx(f.bytes(),options);
    try {const node=leaf(session);if(expected===undefined){assert.equal(node.text.content,'SmartArt');assert.equal(session.report.diagramPreviews[0].status,'partial-cache');}else{assert.equal(node.shape,'parallelogram');assert.equal(node.slant,expected);assert.equal(node.text.content[0].runs[0].text,'Cached text');}} finally {session.dispose();}
  }
});
