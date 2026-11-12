import assert from "node:assert/strict";
import test from "node:test";
import { createPresentation, normalizePresentation } from "@pptkit/core";
import { resolveLayout, resolveNormalizedLayout, resolveShapeTextBox } from "../dist/index.js";

const box = { x: 10, y: 20, width: 100, height: 80 };
function documentWith(input = {}) {
  const document = createPresentation({ size: { width: 200, height: 200 } });
  document.addSlide({ id: "slide" }).addElement({ type: "text", id: "text", content: "Measured", box, frame: { margin: 10 }, ...input });
  return document;
}
const measure = () => ({ bounds: { x: 0, y: 0, width: 40, height: 20 } });
const options = { textMeasurement: { measure }, diagnostics: { textOverflow: true, slideBounds: true } };

test('parallelogram measurement uses its preset text rectangle and margins', () => {
  const d = createPresentation({size:{width:400,height:400}}), s = d.addSlide({id:'s'});
  s.addElement({id:'p',type:'shape',shape:'parallelogram',slant:.4,box:{x:10,y:20,width:100,height:200},text:{content:'Example',frame:{margin:5,wrap:false}}});
  let width;
  const result=resolveLayout(d,{textMeasurement:{measure(request){width=request.availableWidth;return {bounds:{x:0,y:0,width:45,height:20}}}},diagnostics:{textOverflow:true}});
  assert.equal(width,40);
  assert.deepEqual(result.analysis.measurements[0].contentBox,{x:30,y:55,width:40,height:90});
  const warning=result.analysis.diagnostics.find(x=>x.code==='text-overflow');
  assert.equal(warning.elementId,'p');
  assert.deepEqual(warning.available,{x:30,y:55,width:40,height:90});
  assert.deepEqual(result.slides[0].elements[0].box,{x:10,y:20,width:100,height:200});
});

test('shape text rectangle preserves non-parallelogram frames and handles degenerate slants', () => {
  const box={x:10,y:20,width:100,height:100};
  assert.deepEqual(resolveShapeTextBox({shape:'rect',box}),box);
  assert.notEqual(resolveShapeTextBox({shape:'rect',box}),box);
  assert.deepEqual(resolveShapeTextBox({shape:'parallelogram',slant:1,box}),{x:60,y:70,width:0,height:0});
  assert.deepEqual(resolveShapeTextBox({shape:'parallelogram',box:{x:1,y:2,width:0,height:10}}),{x:1,y:2,width:0,height:10});
  const d=createPresentation({size:{width:200,height:200}}),s=d.addSlide({id:'s'});
  s.addElement({id:'p',type:'shape',shape:'parallelogram',slant:1,box,text:{content:'X',frame:{margin:0}}});
  const result=resolveLayout(d,{textMeasurement:{measure:()=>({bounds:{x:0,y:0,width:1,height:1}})},diagnostics:{textOverflow:true}});
  assert.deepEqual(result.analysis.measurements[0].contentBox,{x:50,y:50,width:0,height:0});
  assert.ok(result.analysis.diagnostics.some(x=>x.code==='text-overflow'&&x.elementId==='p'));
});

test('slide-bound checks transform measured parallelogram text about the shape frame', () => {
  const make=rotation=>{const d=createPresentation({size:{width:200,height:200}});d.addSlide({id:'s'}).addElement({id:'p',type:'shape',shape:'parallelogram',slant:.4,box:{x:80,y:20,width:100,height:100},transform:{rotation},text:{content:'X',frame:{margin:0}}});return d};
  const options={textMeasurement:{measure:()=>({bounds:{x:0,y:0,width:100,height:10}})},diagnostics:{slideBounds:true}};
  const outside=resolveLayout(make(0),options).analysis.diagnostics.find(x=>x.code==='slide-overflow'&&x.path.endsWith('.text.content'));
  assert.deepEqual(outside.bounds,{x:105,y:45,width:100,height:10});
  assert.equal(resolveLayout(make(90),options).analysis.diagnostics.some(x=>x.code==='slide-overflow'),false);
});

test('rotated table text maps horizontal measurements into physical cell coordinates', () => {
  for(const [direction,verticalAlign,expected] of [['rotate90','top',{x:84,y:15,width:8,height:30}],['rotate270','top',{x:8,y:35,width:8,height:30}],['rotate90','bottom',{x:5,y:15,width:8,height:30}],['rotate270','bottom',{x:87,y:35,width:8,height:30}]]){
    const d=createPresentation({size:{width:200,height:200}});d.addSlide({id:'s'}).addElement({id:'t',type:'table',box:{x:10,y:20,width:100,height:80},columns:[100],rows:[{height:80,cells:[{content:'text',style:{textDirection:direction,verticalAlign,margin:{left:5,right:5,top:10,bottom:10}}}]}]});
    let request;const result=resolveLayout(d,{textMeasurement:{measure:r=>{request=r;return {bounds:{x:5,y:3,width:30,height:8}}}},diagnostics:{textOverflow:true}});
    assert.equal(request.availableWidth,60);assert.deepEqual(result.analysis.measurements[0].contentBox,{x:5,y:10,width:90,height:60});assert.deepEqual(result.analysis.measurements[0].bounds,expected);
    assert.ok(result.analysis.diagnostics.some(x=>x.code==='table-text-direction-measurement-approximate'));assert.ok(!result.analysis.diagnostics.some(x=>x.code==='text-overflow'));
  }
});

test('rotated table measurements diagnose overflow and preserve missing-provider behavior', () => {
  const d=createPresentation({size:{width:200,height:200}});d.addSlide({id:'s'}).addElement({id:'t',type:'table',box:{x:0,y:0,width:100,height:80},columns:[100],rows:[{height:80,cells:[{content:'text',style:{textDirection:'rotate90',margin:0}}]}]});
  const missing=resolveLayout(d,{diagnostics:{textOverflow:true}});assert.equal(missing.analysis.diagnostics[0].code,'text-measurement-unavailable');assert.equal(missing.analysis.measurements.length,0);
  const result=resolveLayout(d,{textMeasurement:{measure:()=>({bounds:{x:-2,y:0,width:90,height:10}})},diagnostics:{textOverflow:true}});
  assert.ok(result.analysis.diagnostics.some(x=>x.code==='text-overflow'));assert.ok(result.analysis.diagnostics.some(x=>x.code==='table-text-direction-measurement-approximate'));
});

test('opt-in shadow bounds report approximate overflow without changing frame checks',()=>{
 const d=createPresentation({size:{width:100,height:100}}),slide=d.addSlide({id:'s'});slide.addElement({id:'shadowed',type:'shape',shape:'rect',box:{x:80,y:30,width:10,height:10},style:{stroke:{paint:{type:'none'},width:0},shadow:{color:'000000',blur:0,distance:15,direction:0}}});
 assert.equal(resolveLayout(d,{diagnostics:{slideBounds:true}}).analysis.diagnostics.length,0);
 const warning=resolveLayout(d,{diagnostics:{shadowBounds:true}}).analysis.diagnostics.find(x=>x.code==='shadow-slide-overflow-approximate');assert.ok(warning);assert.equal(warning.elementId,'shadowed');assert.deepEqual(warning.bounds,{x:95,y:30,width:10,height:10});assert.deepEqual(warning.available,{x:0,y:0,width:100,height:100});
 assert.equal(resolveLayout(d,{diagnostics:{shadowBounds:true,allowedOffSlideElementIds:['shadowed']}}).analysis.diagnostics.length,0);
 assert.throws(()=>resolveLayout(d,{diagnostics:{shadowBounds:'yes'}}),TypeError);
});

test('shadow estimate transforms group coordinates, rotation and blur and skips invisible ancestors',()=>{
 const make=(opacity=1)=>{const d=createPresentation({size:{width:100,height:100}});d.addSlide({id:'s'}).addElement({id:'g',type:'group',box:{x:50,y:0,width:100,height:100},coordinateSize:{width:50,height:50},opacity,children:[{id:'child',type:'shape',shape:'rect',box:{x:15,y:10,width:10,height:10},transform:{rotation:90},style:{stroke:{paint:{type:'none'},width:0},shadow:{color:'000000',blur:2,distance:10,direction:0}}}]});return d};
 const result=resolveLayout(make(),{diagnostics:{shadowBounds:true}}),warning=result.analysis.diagnostics.find(x=>x.code==='shadow-slide-overflow-approximate');assert.ok(warning);assert.equal(warning.elementId,'child');assert.ok(Math.abs(warning.bounds.x-74)<1e-8);assert.ok(Math.abs(warning.bounds.y-34)<1e-8);assert.ok(Math.abs(warning.bounds.width-32)<1e-8);assert.ok(Math.abs(warning.bounds.height-32)<1e-8);
 assert.equal(resolveLayout(make(0),{diagnostics:{shadowBounds:true}}).analysis.diagnostics.length,0);
 assert.equal(resolveLayout(make(),{diagnostics:{shadowBounds:true,allowedOffSlideElementIds:['g']}}).analysis.diagnostics.length,0);
});

test('shadow estimates skip hidden and zero-opacity effects and honor tolerance',()=>{
 const make=(shadowOpacity=1,hidden=false)=>{const d=createPresentation({size:{width:100,height:100}});d.addSlide({id:'s'}).addElement({id:'shape',type:'shape',shape:'rect',hidden,box:{x:80,y:30,width:10,height:10},style:{stroke:{paint:{type:'none'},width:0},shadow:{color:'000000',opacity:shadowOpacity,blur:0,distance:10.5,direction:0}}});return d};
 assert.equal(resolveLayout(make(0),{diagnostics:{shadowBounds:true}}).analysis.diagnostics.length,0);
 assert.equal(resolveLayout(make(1,true),{diagnostics:{shadowBounds:true}}).analysis.diagnostics.length,0);
 assert.equal(resolveLayout(make(1,true),{diagnostics:{shadowBounds:true,includeHidden:true}}).analysis.diagnostics[0].code,'shadow-slide-overflow-approximate');
 assert.equal(resolveLayout(make(),{diagnostics:{shadowBounds:true,tolerance:1}}).analysis.diagnostics.length,0);
});

test("default resolution remains byte-for-byte JSON compatible and does not add analysis", () => {
  const document = documentWith();
  const plain = resolveLayout(document);
  assert.equal("analysis" in plain, false);
  assert.deepEqual(resolveLayout(document, {}), plain);
  const analyzed = resolveLayout(document, options);
  const { analysis, ...geometry } = analyzed;
  assert.deepEqual(geometry, plain);
  assert.deepEqual(analysis.diagnostics, []);
  assert.equal(analysis.measurements.length, 1);
});

test("providers receive frozen detached runs with theme fonts, effective shrink size and content width", () => {
  const document = createPresentation({ theme: { fonts: { heading: "Header Font", body: "Body Font" } } });
  document.addSlide({ id: "slide" }).addElement({ type: "text", id: "mixed", box, frame: { margin: { left: 12, right: 8, top: 5, bottom: 15 }, wrap: false, autoFit: { mode: "shrink", fontScale: 0.5 } }, content: [
    { style: { indent: 6, hanging: 2, spaceBefore: 3, spaceAfter: 4, bullet: { type: "number", startAt: 2 } }, runs: [
      { text: "Latin ", style: { fontFamily: { theme: "heading" }, fontSize: 24, bold: true } },
      { text: "中文\nnext", style: { fontFamily: { theme: "body" }, fontSize: 20, italic: true, language: "zh-CN" } },
    ] },
  ] });
  let captured;
  const fonts = [];
  const normalized = normalizePresentation(document);
  const before = structuredClone(normalized);
  const result = resolveNormalizedLayout(normalized, { textMeasurement: {
    resolveFont: (font) => { assert.ok(Object.isFrozen(font)); fonts.push(font); return { family: font.family }; },
    measure(request) { captured = request; assert.ok(Object.isFrozen(request.paragraphs[0].runs[0].style)); return measure(); },
  } });
  assert.equal(captured.availableWidth, 80);
  assert.equal(captured.wrap, false);
  assert.deepEqual(fonts.map(({ family, fontSize, text }) => ({ family, fontSize, text })), [
    { family: "Header Font", fontSize: 12, text: "Latin " }, { family: "Body Font", fontSize: 10, text: "中文\nnext" },
  ]);
  assert.equal(captured.paragraphs[0].style.bullet.startAt, 2);
  assert.equal(captured.paragraphs[0].style.hanging, 2);
  assert.deepEqual(result.analysis.measurements[0].contentBox, { x: 12, y: 5, width: 80, height: 60 });
  assert.deepEqual(normalized, before);
});

test("vertical alignment affects measurements but never changes authored geometry", () => {
  for (const [verticalAlign, expected] of [["top", 10], ["middle", 30], ["bottom", 50]]) {
    const result = resolveLayout(documentWith({ frame: { margin: 10, verticalAlign } }), options);
    assert.equal(result.analysis.measurements[0].bounds.y, expected);
    assert.deepEqual(result.slides[0].elements[0].box, box);
  }
});

test("font fallback is explicit, unavailable/failed fonts skip measurement with actionable paths", () => {
  let calls = 0;
  const counted = () => { calls++; return measure(); };
  const fallback = resolveLayout(documentWith(), { textMeasurement: { measure: counted, resolveFont: () => ({ family: "Fallback" }) } });
  assert.equal(calls, 1);
  assert.equal(fallback.analysis.diagnostics[0].code, "font-fallback");
  assert.equal(fallback.analysis.diagnostics[0].path, "slides.0.elements.0.content.0.runs.0.style.fontFamily");
  for (const [resolveFont, code] of [[() => undefined, "font-unavailable"], [() => ({ family: "" }), "font-unavailable"], [() => { throw new Error("font engine offline"); }, "font-resolution-failed"]]) {
    const result = resolveLayout(documentWith(), { textMeasurement: { measure: counted, resolveFont } });
    assert.equal(result.analysis.diagnostics[0].code, code);
    assert.equal(result.analysis.measurements.length, 0);
  }
  assert.equal(calls, 1);
});

test("a missing measurement provider never reports unmeasured text as fitting", () => {
  const result = resolveLayout(documentWith(), { diagnostics: { textOverflow: true } });
  assert.deepEqual(result.analysis.measurements, []);
  assert.equal(result.analysis.diagnostics[0].code, "text-measurement-unavailable");
  assert.equal(result.analysis.diagnostics[0].elementId, "text");
});

test("provider exceptions and malformed bounds produce warnings while preserving input", () => {
  const document = documentWith();
  for (const bounds of [undefined, null, {}, { x: 0, y: 0, width: -1, height: 1 }, { x: NaN, y: 0, width: 1, height: 1 }, { x: 0, y: Infinity, width: 1, height: 1 }, { x: Number.MAX_VALUE, y: 0, width: Number.MAX_VALUE, height: 1 }]) {
    const result = resolveLayout(document, { textMeasurement: { measure: () => ({ bounds }) } });
    assert.equal(result.analysis.diagnostics[0].code, "text-measurement-invalid");
    assert.equal(result.analysis.measurements.length, 0);
    assert.deepEqual(result.slides[0].elements[0].box, box);
  }
  const result = resolveLayout(document, { textMeasurement: { measure: () => { throw new Error("engine unavailable"); } } });
  assert.equal(result.analysis.diagnostics[0].code, "text-measurement-failed");
  assert.match(result.analysis.diagnostics[0].message, /engine unavailable/);
});

test("provider mutation cannot change input and reused provider results are detached", () => {
  const normalized = normalizePresentation(documentWith());
  const before = structuredClone(normalized);
  const mutation = resolveNormalizedLayout(normalized, { textMeasurement: { measure: (request) => { request.paragraphs[0].runs[0].style.fontSize = 999; return measure(); } } });
  assert.equal(mutation.analysis.diagnostics[0].code, "text-measurement-failed");
  const shared = measure();
  const result = resolveNormalizedLayout(normalized, { textMeasurement: { measure: () => shared } });
  shared.bounds.width = 999;
  assert.equal(result.analysis.measurements[0].bounds.width, 40);
  assert.deepEqual(normalized, before);
});

test("text overflow includes left/top and right/bottom excursions and honors tolerance", () => {
  for (const bounds of [{ x: -1, y: 0, width: 1, height: 1 }, { x: 0, y: -1, width: 1, height: 1 }, { x: 0, y: 0, width: 81, height: 1 }, { x: 0, y: 0, width: 1, height: 61 }]) {
    const result = resolveLayout(documentWith(), { textMeasurement: { measure: () => ({ bounds }) }, diagnostics: { textOverflow: true } });
    const warning = result.analysis.diagnostics.find(({ code }) => code === "text-overflow");
    assert.ok(warning);
    assert.equal(warning.path, "slides.0.elements.0.content");
    assert.deepEqual(warning.available, { x: 10, y: 10, width: 80, height: 60 });
  }
  const result = resolveLayout(documentWith(), { textMeasurement: { measure: () => ({ bounds: { x: -0.005, y: 0, width: 80, height: 60 } }) }, diagnostics: { textOverflow: true } });
  assert.deepEqual(result.analysis.diagnostics, []);
});

test("empty content area is measured at zero width instead of inventing usable space", () => {
  let width;
  const result = resolveLayout(documentWith({ frame: { margin: 100 } }), { textMeasurement: { measure: (request) => { width = request.availableWidth; return measure(); } }, diagnostics: { textOverflow: true } });
  assert.equal(width, 0);
  assert.equal(result.analysis.diagnostics[0].code, "text-overflow");
});

test("rotated bounds and transformed nested groups use slide coordinates", () => {
  const document = createPresentation({ size: { width: 100, height: 100 } });
  const slide = document.addSlide({ id: "slide" });
  slide.addElement({ type: "shape", id: "rotated", shape: "rect", box: { x: 10, y: 10, width: 80, height: 10 }, transform: { rotation: 90 } });
  slide.addElement({ type: "group", id: "outer", box: { x: 20, y: 30, width: 200, height: 100 }, coordinateSize: { width: 100, height: 100 }, transform: { rotation: 90 }, children: [
    { type: "group", id: "inner", box: { x: 10, y: 10, width: 50, height: 50 }, coordinateSize: { width: 100, height: 100 }, children: [
      { type: "shape", id: "nested", shape: "rect", box: { x: 90, y: 90, width: 20, height: 20 } },
    ] },
  ] });
  const warnings = resolveLayout(document, { diagnostics: { slideBounds: true } }).analysis.diagnostics;
  const rounded = (id) => Object.fromEntries(Object.entries(warnings.find(({ elementId }) => elementId === id).bounds).map(([key, value]) => [key, Math.round(value)]));
  assert.deepEqual(rounded("rotated"), { x: 45, y: -25, width: 10, height: 80 });
  assert.deepEqual(rounded("nested"), { x: 105, y: 90, width: 10, height: 20 });
  assert.equal(warnings.find(({ elementId }) => elementId === "nested").path, "slides.0.elements.1.children.0.children.0.box");
});

test("flipped groups map child bounds and off-slide exceptions affect only boundary checks", () => {
  const document = createPresentation({ size: { width: 200, height: 200 } });
  document.addSlide().addElement({ type: "group", id: "group", box: { x: 0, y: 0, width: 100, height: 100 }, coordinateSize: { width: 100, height: 100 }, transform: { flipH: true }, children: [
    { type: "text", id: "child", content: "Too wide", box: { x: 120, y: 20, width: 20, height: 20 }, frame: { margin: 0 } },
  ] });
  const result = resolveLayout(document, options);
  assert.deepEqual(result.analysis.diagnostics.find(({ path }) => path.endsWith(".box")).bounds, { x: -40, y: 20, width: 20, height: 20 });
  const allowed = resolveLayout(document, { ...options, diagnostics: { ...options.diagnostics, allowedOffSlideElementIds: ["group"] } });
  assert.equal(allowed.analysis.diagnostics.some(({ code }) => code === "slide-overflow"), false);
  assert.equal(allowed.analysis.diagnostics.some(({ code }) => code === "text-overflow"), true);
});

test("layout elements are measured once with layout identity; hidden subtrees are opt-in", () => {
  const document = createPresentation();
  document.defineSlideLayout({ id: "template", name: "Template", elements: [{ type: "shape", id: "label", shape: "rect", text: { content: "Title" }, box }] });
  document.addSlide({ layoutId: "template" });
  document.addSlide({ layoutId: "template", hidden: true, elements: [{ type: "text", content: "Hidden slide", box }] });
  document.addSlide().addElement({ type: "group", hidden: true, box, coordinateSize: { width: 100, height: 80 }, children: [{ type: "text", content: "Hidden group", box }] });
  const normal = resolveLayout(document, options);
  assert.equal(normal.analysis.measurements.length, 1);
  assert.equal(normal.analysis.measurements[0].layoutId, "template");
  assert.equal(normal.analysis.measurements[0].slideId, undefined);
  assert.match(normal.analysis.measurements[0].path, /^layouts\.1\.elements\.0\.text\.content$/);
  assert.equal(resolveLayout(document, { ...options, diagnostics: { includeHidden: true } }).analysis.measurements.length, 3);
});

test("table cell measurement honors spans and reports grid/frame mismatches", () => {
  const document = createPresentation();
  const slide = document.addSlide();
  slide.addElement({ type: "table", id: "table", box: { x: 0, y: 0, width: 100, height: 40 }, columns: [50, 50], rows: [
    { height: 20, cells: [{ content: "Tall", rowSpan: 2, style: { margin: 0 } }, { content: "Top", style: { margin: 0 } }] },
    { height: 20, cells: [{ content: "Bottom", style: { margin: 0 } }] },
  ] });
  const result = resolveLayout(document, options);
  assert.deepEqual(result.analysis.measurements.map(({ contentBox }) => contentBox), [
    { x: 0, y: 0, width: 50, height: 40 }, { x: 50, y: 0, width: 50, height: 20 }, { x: 50, y: 20, width: 50, height: 20 },
  ]);
  assert.deepEqual(result.analysis.diagnostics, []);
  const normalized = normalizePresentation(document);
  normalized.slides[0].elements[0].box.height = 100;
  assert.equal(resolveNormalizedLayout(normalized, options).analysis.diagnostics[0].code, "table-geometry-mismatch");
  normalized.slides[0].elements[0].rows[1].cells[0].rowSpan = 3;
  assert.ok(resolveNormalizedLayout(normalized, options).analysis.diagnostics.some(({ code }) => code === "table-span-invalid"));
});

test("measured glyph bounds can exceed the slide even when their text frame fits", () => {
  const result = resolveLayout(documentWith({ box: { x: 170, y: 10, width: 20, height: 30 }, frame: { margin: 0 } }), options);
  const warning = result.analysis.diagnostics.find(({ code }) => code === "slide-overflow");
  assert.equal(warning.path, "slides.0.elements.0.content");
  assert.deepEqual(warning.bounds, { x: 170, y: 10, width: 40, height: 20 });
});

test("invalid diagnostic configuration fails explicitly", () => {
  for (const configuration of [
    { textMeasurement: {} }, { textMeasurement: { measure: null } }, { textMeasurement: { measure, resolveFont: 3 } },
    { diagnostics: { tolerance: -1 } }, { diagnostics: { tolerance: NaN } }, { diagnostics: { slideBounds: "true" } },
    { diagnostics: { allowedOffSlideElementIds: [2] } },
  ]) assert.throws(() => resolveLayout(documentWith(), configuration));
});

test("repeated resolution with a deterministic provider returns identical detached analysis", () => {
  const document = documentWith();
  const first = resolveLayout(document, options);
  const second = resolveLayout(document, options);
  assert.deepEqual(first, second);
  first.analysis.measurements[0].bounds.x = 999;
  assert.notEqual(first.analysis.measurements[0].bounds.x, second.analysis.measurements[0].bounds.x);
});


test("accidentally asynchronous providers fail explicitly without unhandled rejections", async () => {
  for (const measure of [async () => ({ bounds: box }), async () => { throw new Error("late failure"); }]) {
    const result = resolveLayout(documentWith(), { textMeasurement: { measure } });
    assert.equal(result.analysis.diagnostics[0].code, "text-measurement-failed");
    assert.match(result.analysis.diagnostics[0].message, /synchronous/);
  }
  const result = resolveLayout(documentWith(), { textMeasurement: { measure, resolveFont: async () => { throw new Error("font load failed"); } } });
  assert.equal(result.analysis.diagnostics[0].code, "font-resolution-failed");
  await new Promise((resolve) => setImmediate(resolve));
});

test("extreme nested scale cannot silently pass slide-boundary analysis", () => {
  const normalized = normalizePresentation(documentWith());
  normalized.slides[0].elements[0].box = { x: Number.MAX_VALUE, y: 0, width: Number.MAX_VALUE, height: 10 };
  const result = resolveNormalizedLayout(normalized, { diagnostics: { slideBounds: true } });
  assert.equal(result.analysis.diagnostics[0].code, "geometry-unmeasurable");
  assert.equal(result.analysis.diagnostics[0].bounds, undefined);
});

test('East Asian vertical provider uses frozen logical axes and maps alignment from the right edge', () => {
  for (const [verticalAlign, x] of [["top", 70], ["middle", 37.5], ["bottom", 5]]) {
    const deck = documentWith({ content: "日本語ABC", frame: {
      textDirection: "eastAsianVertical", verticalAlign, margin: { left: 5, right: 7, top: 11, bottom: 13 },
    } });
    let calls = 0;
    const result = resolveLayout(deck, { textMeasurement: {
      measure() { assert.fail("Horizontal callback cannot measure vertical text"); },
      measureEastAsianVertical(request) {
        calls++;
        assert.ok(Object.isFrozen(request) && Object.isFrozen(request.paragraphs[0].runs));
        assert.equal(request.availableInlineSize, 56);
        assert.equal(request.availableBlockSize, 88);
        assert.equal(request.paragraphs[0].runs.map(run => run.text).join(""), "日本語ABC");
        return { bounds: { x: 2, y: 3, width: 40, height: 20 } };
      },
    }, diagnostics: { textOverflow: true } });
    assert.equal(calls, 1);
    assert.deepEqual(result.analysis.diagnostics, []);
    assert.deepEqual(result.analysis.measurements[0].bounds, { x, y: 13, width: 20, height: 40 });
  }
});

test('East Asian vertical bounds diagnose overflow and reject invalid provider results', () => {
  const deck = documentWith({ frame: { textDirection: "eastAsianVertical", margin: 10 } });
  for (const [width, code, count] of [[61, "text-overflow", 1], [Infinity, "text-measurement-invalid", 0]]) {
    const result = resolveLayout(deck, { textMeasurement: { measure, measureEastAsianVertical: () => ({
      bounds: { x: 0, y: 0, width, height: 20 },
    }) }, diagnostics: { textOverflow: true } });
    assert.ok(result.analysis.diagnostics.some(d => d.code === code));
    assert.equal(result.analysis.measurements.length, count);
  }
});

test('East Asian vertical table analysis reports the physical cell position', () => {
  const deck = createPresentation();
  deck.addSlide().addElement({ id: 'table', type: 'table', box: { x: 40, y: 50, width: 100, height: 80 },
    columns: [30, 70], rows: [{ height: 80, cells: [
      { content: 'control', style: { margin: 0 } },
      { content: '日本語', style: { textDirection: 'eastAsianVertical', margin: 5, verticalAlign: 'top' } },
    ] }] });
  const result = resolveLayout(deck, { textMeasurement: { measure: () => ({ bounds: { x: 0, y: 0, width: 10, height: 10 } }),
    measureEastAsianVertical(request) {
      assert.equal(request.availableInlineSize, 70);
      assert.equal(request.availableBlockSize, 60);
      return { bounds: { x: 1, y: 2, width: 40, height: 16 } };
    } } });
  assert.deepEqual(result.analysis.diagnostics, []);
  const cell = result.analysis.measurements.find(m => m.path.endsWith('.cells.1.content'));
  assert.deepEqual(cell.contentBox, { x: 35, y: 5, width: 60, height: 70 });
  assert.deepEqual(cell.bounds, { x: 77, y: 6, width: 16, height: 40 });
});
