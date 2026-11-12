import assert from "node:assert/strict";
import test from "node:test";
import { createPresentation, normalizePresentation } from "@pptkit/core";
import { resolveLayout, resolveNormalizedLayout, resolveShapeTextBox } from "../dist/index.js";

function textFixture() {
  const document = createPresentation({ size: { width: 300, height: 200 } });
  document.addSlide({ id: "slide" }).addElement({ type: "text", id: "text", content: "abcdefghij", box: { x: 10, y: 20, width: 60, height: 20 }, frame: { margin: 0 } });
  return document;
}
const measure = (request) => {
  const length = request.paragraphs.reduce((sum, p) => sum + p.runs.reduce((n, r) => n + r.text.length, 0), 0);
  const perLine = request.wrap ? Math.max(1, Math.floor(request.availableWidth / 10)) : length;
  return { bounds: { x: 0, y: 0, width: Math.min(length, perLine) * 10, height: Math.ceil(length / Math.max(1, perLine)) * 20 } };
};
const fit = (document, rules, extra = {}) => resolveLayout(document, { textMeasurement: { measure }, intrinsicSizing: rules, ...extra });
const first = (result) => result.slides[0].elements[0];
const content = { type: "content" };

test('parallelogram natural width inverts its preset text rectangle in both aspect regions', () => {
  for(const [slant,characters,expected] of [[.4,3,80],[.4,9,160],[1,3,148],[0,3,48]]) {
    const d=createPresentation({size:{width:400,height:300}});d.addSlide({id:'s'}).addElement({id:'p',type:'shape',shape:'parallelogram',slant,box:{x:10,y:20,width:40,height:100},text:{content:'x'.repeat(characters),frame:{margin:5}}});
    const result=fit(d,[{elementId:'p',width:content}],{diagnostics:{textOverflow:true}});
    assert.ok(Math.abs(first(result).box.width-expected)<1e-9,`slant ${slant}, characters ${characters}`);
    assert.ok(Math.abs(result.analysis.measurements[0].contentBox.width-characters*10)<1e-9);
    const warning=result.analysis.diagnostics.find(x=>x.code==='text-overflow');
    if(warning) assert.ok(warning.bounds.x+warning.bounds.width<=warning.available.x+warning.available.width+1e-9);
  }
});

test('wrapped parallelogram with both automatic axes rejects unbreakable overflow atomically', () => {
  for(const rule of [{elementId:'p',width:{type:'content',max:1},height:content}]) {
    const d=textFixture();d.slides[0].addElement({id:'p',type:'shape',shape:'parallelogram',box:{x:20,y:30,width:80,height:100},text:{content:'shape',frame:{margin:0}}});
    const before=resolveLayout(d).slides[0].elements.map(x=>x.box), result=fit(d,[{elementId:'text',width:content},rule]);
    assert.equal(result.analysis.diagnostics[0].code,'intrinsic-shape-height-unavailable');
    assert.deepEqual(result.slides[0].elements.map(x=>x.box),before);
  }
});

test('wrapped parallelogram content width and height are both fitted and verified', () => {
  const d=createPresentation({size:{width:400,height:300}});d.addSlide({id:'s'}).addElement({id:'p',type:'shape',shape:'parallelogram',slant:.4,box:{x:0,y:0,width:20,height:20},text:{content:'wrapped',frame:{margin:0,wrap:true}}});
  const seen=[];
  const result=resolveLayout(d,{intrinsicSizing:[{elementId:'p',width:{type:'content',max:80},height:content}],textMeasurement:{measure:r=>{seen.push([r.wrap,r.availableWidth]);return {bounds:r.wrap?{x:0,y:0,width:Math.min(30,r.availableWidth),height:r.availableWidth>=63?20:40}:{x:0,y:0,width:100,height:20}}}},diagnostics:{textOverflow:true}});
  assert.equal(first(result).box.width,80);assert.ok(first(result).box.height>=80-1e-8);assert.equal(seen[0][0],false);assert.ok(seen.slice(1).every(x=>x[0]));
  assert.ok(!result.analysis.diagnostics.some(x=>x.code==='text-overflow'||x.code==='intrinsic-shape-height-unavailable'));
});

test('wrapped two-axis sizing widens after horizontal overflow and respects the maximum', () => {
  const make=()=>{const d=createPresentation({size:{width:400,height:300}});d.addSlide({id:'s'}).addElement({id:'p',type:'shape',shape:'parallelogram',slant:.4,box:{x:0,y:0,width:20,height:20},text:{content:'wrapped',frame:{margin:0,wrap:true}}});return d};
  const seen=[],measure=r=>{seen.push([r.wrap,r.availableWidth]);return {bounds:{x:0,y:0,width:r.wrap?100:10,height:20}}};
  const options=max=>({intrinsicSizing:[{elementId:'p',width:{type:'content',max},height:content}],textMeasurement:{measure},diagnostics:{textOverflow:true}});
  const result=resolveLayout(make(),options(300));assert.ok(first(result).box.width>120);assert.ok(!result.analysis.diagnostics.some(x=>x.code==='text-overflow'||x.code==='intrinsic-shape-height-unavailable'));assert.ok(seen.filter(x=>x[0]).length>=4);
  const failed=resolveLayout(make(),options(50));assert.equal(failed.analysis.diagnostics[0].code,'intrinsic-shape-height-unavailable');assert.deepEqual(first(failed).box,{x:0,y:0,width:20,height:20});
  let calls=0;
  const unstable=resolveLayout(make(),{intrinsicSizing:[{elementId:'p',width:content,height:content}],textMeasurement:{measure:r=>{calls++;return {bounds:{x:0,y:0,width:r.wrap?r.availableWidth+100:10,height:20}}}}});
  assert.equal(unstable.analysis.diagnostics[0].code,'intrinsic-shape-height-unavailable');assert.match(unstable.analysis.diagnostics[0].message,/16 iterations/);assert.ok(calls<=50);assert.deepEqual(first(unstable).box,{x:0,y:0,width:20,height:20});
});

test('wrapped two-axis sizing retries height-bound failures without hiding provider errors', () => {
  const make=()=>{const d=createPresentation({size:{width:400,height:300}});d.addSlide({id:'s'}).addElement({id:'p',type:'shape',shape:'parallelogram',slant:.4,box:{x:0,y:0,width:20,height:20},text:{content:'wrapped',frame:{margin:0,wrap:true}}});return d};
  const rule=max=>({elementId:'p',width:{type:'content',max},height:{type:'content',max:60}});
  const measure=r=>({bounds:{x:0,y:0,width:r.wrap?Math.min(10,r.availableWidth):10,height:r.wrap&&r.availableWidth<50?80:20}});
  const result=resolveLayout(make(),{intrinsicSizing:[rule(300)],textMeasurement:{measure},diagnostics:{textOverflow:true}});
  assert.ok(first(result).box.width>=80);assert.ok(first(result).box.height<=60);assert.deepEqual(result.analysis.diagnostics,[]);
  const capped=resolveLayout(make(),{intrinsicSizing:[rule(30)],textMeasurement:{measure}});
  assert.equal(capped.analysis.diagnostics[0].code,'intrinsic-shape-height-unavailable');assert.deepEqual(first(capped).box,{x:0,y:0,width:20,height:20});
  let wraps=0;
  const broken=resolveLayout(make(),{intrinsicSizing:[rule(300)],textMeasurement:{measure:r=>{if(r.wrap){wraps++;throw new Error('provider failed');}return {bounds:{x:0,y:0,width:10,height:20}};}}});
  assert.equal(broken.analysis.diagnostics[0].code,'text-measurement-failed');assert.equal(wraps,2);assert.deepEqual(first(broken).box,{x:0,y:0,width:20,height:20});
});

test('wrapped parallelogram height iterates on the actual text width and verifies the fit', () => {
  const d=createPresentation({size:{width:400,height:300}});d.addSlide({id:'s'}).addElement({id:'p',type:'shape',shape:'parallelogram',slant:.4,box:{x:0,y:0,width:100,height:20},text:{content:'wrapped',frame:{margin:0,wrap:true}}});
  const widths=[];
  const result=resolveLayout(d,{intrinsicSizing:[{elementId:'p',height:content}],textMeasurement:{measure:r=>{widths.push(r.availableWidth);return {bounds:{x:0,y:0,width:10,height:r.availableWidth>=80?20:40}}}},diagnostics:{textOverflow:true}});
  assert.ok(widths[0]>80);assert.ok(widths[1]<80);assert.ok(first(result).box.height>20);
  assert.ok(!result.analysis.diagnostics.some(x=>x.code==='text-overflow'||x.code==='intrinsic-shape-height-unavailable'));
});

test('wrapped parallelogram sizing bounds unstable providers and rolls back on failure', () => {
  const make=()=>{const d=createPresentation({size:{width:400,height:300}});d.addSlide({id:'s'}).addElement({id:'p',type:'shape',shape:'parallelogram',slant:0,box:{x:0,y:0,width:100,height:20},text:{content:'wrapped',frame:{margin:0,wrap:true}}});return d};
  let calls=0;
  const unstable=resolveLayout(make(),{intrinsicSizing:[{elementId:'p',height:content}],textMeasurement:{measure:()=>({bounds:{x:0,y:0,width:1,height:10+calls++}})}});
  assert.equal(unstable.analysis.diagnostics[0].code,'intrinsic-shape-height-unavailable');assert.match(unstable.analysis.diagnostics[0].message,/32/);assert.ok(calls<=34);assert.equal(first(unstable).box.height,20);
  const capped=resolveLayout(make(),{intrinsicSizing:[{elementId:'p',height:{type:'content',max:10}}],textMeasurement:{measure:()=>({bounds:{x:0,y:0,width:1,height:20}})}});
  assert.equal(capped.analysis.diagnostics[0].code,'intrinsic-shape-height-unavailable');assert.equal(first(capped).box.height,20);
});

test('parallelogram width uses the resolved parent height and respects maximum bounds', () => {
  const d=createPresentation({size:{width:400,height:300}});d.addSlide({id:'s'}).addElement({id:'p',type:'shape',shape:'parallelogram',slant:.4,box:{x:0,y:0,width:20,height:20},text:{content:'xxx',frame:{margin:5,wrap:false}}});
  const result=fit(d,[{elementId:'p',width:{type:'content',max:60},height:{type:'parent',fraction:.5}}],{diagnostics:{textOverflow:true}});
  assert.equal(first(result).box.height,150);
  assert.equal(first(result).box.width,60);
  assert.ok(result.analysis.diagnostics.some(x=>x.code==='text-overflow'&&x.elementId==='p'));
});

test('unwrapped parallelogram height fits the preset rectangle in both height regions', () => {
  for(const [slant,measuredHeight,expected] of [[0,20,36],[.4,20,43.60589701950147],[.4,60,140],[1,0,13.944487245360106]]) {
    const d=createPresentation({size:{width:400,height:300}});d.addSlide({id:'s'}).addElement({id:'p',type:'shape',shape:'parallelogram',slant,box:{x:10,y:20,width:100,height:20},text:{content:'x',frame:{margin:5,wrap:false}}});
    const result=resolveLayout(d,{intrinsicSizing:[{elementId:'p',height:content}],textMeasurement:{measure:()=>({bounds:{x:0,y:0,width:1,height:measuredHeight}})},diagnostics:{textOverflow:true}});
    assert.ok(Math.abs(first(result).box.height-expected)<1e-8,`slant ${slant}: ${first(result).box.height}`);
    assert.ok(!result.analysis.diagnostics.some(x=>x.code==='text-overflow'));
  }
});

test('unwrapped parallelogram height rejects impossible or width-dependent measurements atomically', () => {
  const make=slant=>{const d=createPresentation({size:{width:400,height:300}});d.addSlide({id:'s'}).addElement({id:'p',type:'shape',shape:'parallelogram',slant,box:{x:0,y:0,width:100,height:20},text:{content:'x',frame:{margin:0,wrap:false}}});return d};
  for(const [d,provider,code,rule] of [
    [make(1),()=>({bounds:{x:0,y:0,width:1,height:30}}),'intrinsic-size-unsatisfiable',{elementId:'p',height:content}],
    [make(.4),r=>({bounds:{x:0,y:0,width:1,height:r.availableWidth?80:10}}),'intrinsic-shape-height-unavailable',{elementId:'p',height:content}],
    [make(.4),()=>({bounds:{x:0,y:0,width:1,height:30}}),'intrinsic-size-unsatisfiable',{elementId:'p',height:{type:'content',max:20}}],
  ]) {
    const result=resolveLayout(d,{intrinsicSizing:[rule],textMeasurement:{measure:provider}});
    assert.equal(result.analysis.diagnostics[0].code,code);
    assert.equal(first(result).box.height,20);
  }
});

test('both content axes fit unwrapped parallelogram text and remeasure after sizing', () => {
  for(const [slant,width,height] of [[.4,480/7,360/7],[1,144,72]]) {
    const d=createPresentation({size:{width:400,height:300}});d.addSlide({id:'s'}).addElement({id:'p',type:'shape',shape:'parallelogram',slant,box:{x:0,y:0,width:20,height:20},text:{content:'x',frame:{margin:5,wrap:false}}});
    const result=resolveLayout(d,{intrinsicSizing:[{elementId:'p',width:content,height:content}],textMeasurement:{measure:()=>({bounds:{x:0,y:0,width:30,height:20}})},diagnostics:{textOverflow:true}});
    assert.ok(Math.abs(first(result).box.width-width)<1e-8);assert.ok(Math.abs(first(result).box.height-height)<1e-8);
    assert.ok(!result.analysis.diagnostics.some(x=>x.code==='text-overflow'));
  }
});

test('both content axes respect minima and reject maximum or changed measurement atomically', () => {
  const make=()=>{const d=createPresentation({size:{width:400,height:300}});d.addSlide({id:'s'}).addElement({id:'p',type:'shape',shape:'parallelogram',slant:.8,box:{x:0,y:0,width:20,height:20},text:{content:'x',frame:{margin:0,wrap:false}}});return d};
  const measure=()=>({bounds:{x:0,y:0,width:30,height:20}});
  const result=resolveLayout(make(),{intrinsicSizing:[{elementId:'p',width:content,height:{type:'content',min:100}}],textMeasurement:{measure},diagnostics:{textOverflow:true}});
  assert.equal(first(result).box.height,100);assert.ok(!result.analysis.diagnostics.some(x=>x.code==='text-overflow'));
  for(const [provider,rule,code] of [[measure,{elementId:'p',width:{type:'content',max:10},height:content},'intrinsic-size-unsatisfiable'],[r=>({bounds:{x:0,y:0,width:r.availableWidth?300:30,height:20}}),{elementId:'p',width:content,height:content},'intrinsic-shape-height-unavailable'],[()=>({bounds:{x:-1,y:0,width:30,height:20}}),{elementId:'p',width:content,height:content},'intrinsic-shape-height-unavailable']]) {
    const failed=resolveLayout(make(),{intrinsicSizing:[rule],textMeasurement:{measure:provider}});
    assert.equal(failed.analysis.diagnostics[0].code,code);assert.deepEqual(first(failed).box,{x:0,y:0,width:20,height:20});
  }
});

test('both content axes remeasure authored paragraph alignment before applying the size', () => {
  const d=createPresentation({size:{width:400,height:300}});d.addSlide({id:'s'}).addElement({id:'p',type:'shape',shape:'parallelogram',box:{x:0,y:0,width:20,height:20},text:{content:[{style:{align:'right'},runs:[{text:'x'}]}],frame:{margin:0,wrap:false}}});
  const seen=[];
  const result=resolveLayout(d,{intrinsicSizing:[{elementId:'p',width:content,height:content}],textMeasurement:{measure:r=>{seen.push(r.paragraphs[0].style.align);return {bounds:{x:0,y:0,width:r.paragraphs[0].style.align==='right'?300:30,height:20}}}}});
  assert.deepEqual(seen.slice(0,2),['left','right']);assert.equal(result.analysis.diagnostics[0].code,'intrinsic-shape-height-unavailable');assert.equal(first(result).box.width,20);
});

test('both content axes retry a shorter maximum height with a wider frame', () => {
  const make=slant=>{const d=createPresentation({size:{width:400,height:300}});d.addSlide({id:'s'}).addElement({id:'p',type:'shape',shape:'parallelogram',slant,box:{x:0,y:0,width:20,height:20},text:{content:'x',frame:{margin:5,wrap:false}}});return d};
  const measure=()=>({bounds:{x:0,y:0,width:30,height:20}});
  const options=(maxHeight,maxWidth)=>({intrinsicSizing:[{elementId:'p',width:{type:'content',...(maxWidth===undefined?{}:{max:maxWidth})},height:{type:'content',max:maxHeight}}],textMeasurement:{measure},diagnostics:{textOverflow:true}});
  const result=resolveLayout(make(1),options(50));
  assert.equal(first(result).box.height,50);assert.ok(Math.abs(first(result).box.width-1250/7)<1e-8);
  assert.ok(!result.analysis.diagnostics.some(x=>x.code==='text-overflow'));
  const tall=resolveLayout(make(0),{intrinsicSizing:[{elementId:'p',width:{type:'content',max:60},height:{type:'content',min:100}}],textMeasurement:{measure},diagnostics:{textOverflow:true}});
  assert.equal(first(tall).box.height,100);assert.equal(first(tall).box.width,48);assert.ok(!tall.analysis.diagnostics.some(x=>x.code==='text-overflow'));
  for(const failed of [resolveLayout(make(1),options(50,150)),resolveLayout(make(1),options(36)),resolveLayout(make(0),options(35))]) {
    assert.equal(failed.analysis.diagnostics[0].code,'intrinsic-size-unsatisfiable');assert.deepEqual(first(failed).box,{x:0,y:0,width:20,height:20});
  }
});

test('minimum height handles the nonmonotonic preset text rectangle', () => {
  const make=slant=>{const d=createPresentation({size:{width:400,height:300}});d.addSlide({id:'s'}).addElement({id:'p',type:'shape',shape:'parallelogram',slant,box:{x:0,y:0,width:100,height:20},text:{content:'x',frame:{margin:0,wrap:false}}});return d};
  const measure=()=>({bounds:{x:0,y:0,width:1,height:20}});
  const run=(slant,min,max)=>resolveLayout(make(slant),{intrinsicSizing:[{elementId:'p',height:{type:'content',min,...(max===undefined?{}:{max})}}],textMeasurement:{measure},diagnostics:{textOverflow:true}});
  assert.equal(first(run(.8,90)).box.height,90);
  const jumped=run(.8,98);
  assert.ok(Math.abs(first(jumped).box.height-120)<1e-9);
  assert.ok(!jumped.analysis.diagnostics.some(x=>x.code==='text-overflow'));
  const capped=run(.8,98,110);
  assert.equal(capped.analysis.diagnostics[0].code,'intrinsic-size-unsatisfiable');
  assert.equal(first(capped).box.height,20);
  const impossible=run(1,90);
  assert.equal(impossible.analysis.diagnostics[0].code,'intrinsic-size-unsatisfiable');
  assert.equal(first(impossible).box.height,20);
  const enormous=resolveLayout(make(.9),{intrinsicSizing:[{elementId:'p',height:content}],textMeasurement:{measure:()=>({bounds:{x:0,y:0,width:1,height:1e308}})}});
  assert.equal(enormous.analysis.diagnostics[0].code,'intrinsic-size-unsatisfiable');
  assert.equal(first(enormous).box.height,20);
});

test("only explicitly automatic axes change, preserving fixed dimensions and authored state", () => {
  const document = textFixture();
  const before = normalizePresentation(document);
  const result = fit(document, [{ elementId: "text", height: content }]);
  assert.deepEqual(first(result).box, { x: 10, y: 20, width: 60, height: 40 });
  assert.deepEqual(normalizePresentation(document), before);
  assert.deepEqual(first(resolveLayout(document)).box, before.slides[0].elements[0].box);
  assert.deepEqual(result, fit(document, [{ elementId: "text", height: content }]));
});

test("bounded natural text width precedes wrapped height and includes margins", () => {
  const normalized = normalizePresentation(textFixture());
  normalized.slides[0].elements[0].frame.margin = { left: 5, right: 5, top: 5, bottom: 5 };
  normalized.slides[0].elements[0].content[0].style.align = "right";
  const calls = [];
  const result = resolveNormalizedLayout(normalized, { textMeasurement: { measure: (request) => { calls.push(request); return measure(request); } }, intrinsicSizing: [{ elementId: "text", width: { type: "content", max: 60 }, height: content }] });
  assert.deepEqual(first(result).box, { x: 10, y: 20, width: 60, height: 50 });
  assert.equal(calls[0].wrap, false);
  assert.equal(calls[0].paragraphs[0].style.align, "left");
  assert.equal(calls[1].availableWidth, 50);
  assert.equal(calls[1].paragraphs[0].style.align, "right");
});

test("text minimum height is honored and impossible maximum atomically retains all intrinsic rules", () => {
  const document = textFixture();
  document.slides[0].addElement({ type: "text", id: "other", content: "short", box: { x: 0, y: 0, width: 20, height: 20 }, frame: { margin: 0 } });
  assert.equal(first(fit(document, [{ elementId: "text", height: { type: "content", min: 80 } }])).box.height, 80);
  const result = fit(document, [{ elementId: "other", width: content }, { elementId: "text", height: { type: "content", max: 10 } }]);
  assert.equal(result.analysis.diagnostics[0].code, "intrinsic-size-unsatisfiable");
  assert.deepEqual(result.slides[0].elements.map(({ box }) => box), resolveLayout(document).slides[0].elements.map(({ box }) => box));
});

function imageFixture(known = true) {
  const document = createPresentation();
  const asset = document.registerAsset({ kind: "image", id: "asset", source: { type: "url", value: "https://example.invalid/never-load.png" }, ...(known ? { width: 200, height: 100 } : {}) });
  document.addSlide().addElement({ type: "image", id: "image", assetId: asset.id, box: { x: 1, y: 2, width: 80, height: 80 }, fit: "stretch" });
  return document;
}

test("image automatic axes preserve natural aspect ratios with fixed and bounded sizes", () => {
  assert.deepEqual(first(fit(imageFixture(), [{ elementId: "image", height: content }])).box, { x: 1, y: 2, width: 80, height: 40 });
  const both = fit(imageFixture(), [{ elementId: "image", width: { type: "content", max: 100 }, height: { type: "content", min: 30 } }]);
  assert.deepEqual(first(both).box, { x: 1, y: 2, width: 100, height: 50 });
  const large = fit(imageFixture(), [{ elementId: "image", width: { type: "content", min: 400 }, height: content }]);
  assert.deepEqual(first(large).box, { x: 1, y: 2, width: 400, height: 200 });
});

test("missing image dimensions and incompatible aspect bounds produce actionable failure", () => {
  const missing = fit(imageFixture(false), [{ elementId: "image", height: content }]);
  assert.equal(missing.analysis.diagnostics[0].code, "intrinsic-image-size-unavailable");
  for (const rule of [
    { elementId: "image", height: { type: "content", min: 60 } },
    { elementId: "image", width: { type: "content", min: 200 }, height: { type: "content", max: 50 } },
  ]) {
    const result = fit(imageFixture(), [rule]);
    assert.equal(result.analysis.diagnostics[0].code, "intrinsic-size-unsatisfiable");
    assert.equal(first(result).box.width, 80);
  }
});

function groupFixture(transform = {}) {
  const document = createPresentation({ size: { width: 300, height: 200 } });
  document.addSlide({ id: "slide" }).addElement({ type: "group", id: "group", box: { x: 70, y: 50, width: 200, height: 100 }, coordinateSize: { width: 100, height: 100 }, transform, children: [
    { type: "shape", id: "child", shape: "rect", box: { x: -20, y: -10, width: 60, height: 50 } },
  ] });
  return document;
}
function world(group, point) {
  const cx = group.box.x + group.box.width / 2, cy = group.box.y + group.box.height / 2;
  const x = (point.x * group.box.width / group.coordinateSize.width - group.box.width / 2) * (group.transform.flipH ? -1 : 1);
  const y = (point.y * group.box.height / group.coordinateSize.height - group.box.height / 2) * (group.transform.flipV ? -1 : 1);
  const angle = group.transform.rotation * Math.PI / 180;
  return { x: cx + x * Math.cos(angle) - y * Math.sin(angle), y: cy + x * Math.sin(angle) + y * Math.cos(angle) };
}

test("group fitting rebases negative children while preserving world positions and scale under flips/rotation", () => {
  for (const rotation of [0, 25, 90, 180]) for (const flipH of [false, true]) for (const flipV of [false, true]) {
    const document = groupFixture({ rotation, flipH, flipV });
    const before = first(resolveLayout(document));
    const result = fit(document, [{ elementId: "group", width: content, height: content }]);
    const group = first(result);
    assert.deepEqual(result.analysis.diagnostics, []);
    assert.deepEqual(group.coordinateSize, { width: 60, height: 50 });
    assert.equal(group.box.width, 120); assert.equal(group.box.height, 50);
    const child = group.children[0];
    assert.deepEqual(child.box, { x: 0, y: 0, width: 60, height: 50 });
    for (const [dx, dy] of [[0, 0], [60, 0], [0, 50], [60, 50]]) {
      const old = world(before, { x: before.children[0].box.x + dx, y: before.children[0].box.y + dy });
      const actual = world(group, { x: child.box.x + dx, y: child.box.y + dy });
      assert.ok(Math.abs(old.x - actual.x) < 1e-8 && Math.abs(old.y - actual.y) < 1e-8);
    }
  }
});

test("parent-relative child dimensions use group local coordinateSize and slide dimensions", () => {
  const result = fit(groupFixture(), [{ elementId: "child", width: { type: "parent", fraction: 0.5 }, height: { type: "parent", fraction: 0.25 } }]);
  assert.equal(first(result).children[0].box.width, 50);
  assert.equal(first(result).children[0].box.height, 25);
  assert.equal(first(fit(textFixture(), [{ elementId: "text", width: { type: "parent", fraction: 0.5 }, height: content }])).box.width, 150);
});

test("parent/content sizing cycles reject the batch regardless of rule order", () => {
  const rules = [{ elementId: "group", width: content }, { elementId: "child", width: { type: "parent" } }];
  for (const order of [rules, [...rules].reverse()]) {
    const document = groupFixture();
    const result = fit(document, order);
    assert.ok(result.analysis.diagnostics.some(({ code }) => code === "intrinsic-sizing-cycle"));
    assert.deepEqual(first(result), first(resolveLayout(document)));
  }
});

test("fitted groups include free connector routes and updated anchored connector bounds", () => {
  const document = groupFixture();
  const normalized = normalizePresentation(document);
  const group = normalized.slides[0].elements[0];
  const source = createPresentation();
  source.addSlide().addElement({ type: "connector", id: "free", start: { x: -50, y: -20 }, route: [{ x: 80, y: 60 }], end: { x: 100, y: 10 } });
  group.children.push(normalizePresentation(source).slides[0].elements[0]);
  const result = resolveNormalizedLayout(normalized, { intrinsicSizing: [{ elementId: "group", width: content, height: content }] });
  const fitted = first(result);
  assert.deepEqual(fitted.coordinateSize, { width: 150, height: 80 });
  assert.deepEqual(fitted.children[1].start, { x: 0, y: 0 });
  assert.deepEqual(fitted.children[1].route, [{ x: 130, y: 80 }]);
  const connector = group.children[1]; connector.start = { elementId: "child", anchor: "right" };
  assert.deepEqual(resolveNormalizedLayout(normalized, { intrinsicSizing: [{ elementId: "group", width: content }] }).analysis.diagnostics, []);
});

test("hidden subtrees, unsupported targets and unavailable providers are explicit", () => {
  const normalized = normalizePresentation(textFixture()); normalized.slides[0].hidden = true;
  assert.equal(first(resolveNormalizedLayout(normalized, { intrinsicSizing: [{ elementId: "text", height: content }] })).box.height, 20);
  const missingProvider = resolveNormalizedLayout(normalized, { intrinsicSizing: [{ elementId: "text", height: content, includeHidden: true }] });
  assert.equal(missingProvider.analysis.diagnostics[0].code, "text-measurement-unavailable");
  const missing = fit(textFixture(), [{ elementId: "missing", height: content }]);
  assert.equal(missing.analysis.diagnostics[0].path, "options.intrinsicSizing.0");
});

test("group minima add spare area without changing child geometry and maxima fail", () => {
  const document = groupFixture();
  const result = fit(document, [{ elementId: "group", width: { type: "content", min: 180 }, height: content }]);
  assert.equal(first(result).box.width, 180);
  assert.equal(first(result).coordinateSize.width, 90);
  assert.equal(first(result).children[0].box.width, 60);
  assert.equal(fit(document, [{ elementId: "group", width: { type: "content", max: 100 } }]).analysis.diagnostics[0].code, "intrinsic-size-unsatisfiable");
});

test("composition and connector resolution consume intrinsic geometry", () => {
  const document = textFixture();
  document.slides[0].addElement({ type: "connector", id: "link", start: { elementId: "text", anchor: "bottom" }, end: { x: 200, y: 200 } });
  const result = fit(document, [{ elementId: "text", height: content }], { composition: [{ type: "align", container: { kind: "slide", id: "slide" }, elementIds: ["text"], vertical: "end" }] });
  assert.equal(first(result).box.y, 160);
  assert.deepEqual(result.slides[0].elements[1].start, { x: 40, y: 200 });
});

test("invalid dimension contracts reject before any sizing", () => {
  for (const rule of [{ elementId: "text" }, { elementId: "text", width: { type: "content", min: -1 } }, { elementId: "text", width: { type: "content", min: 5, max: 2 } }, { elementId: "text", width: { type: "parent", fraction: Infinity } }, { elementId: "text", height: { type: "auto" } }]) assert.throws(() => fit(textFixture(), [rule]));
  assert.throws(() => fit(textFixture(), [{ elementId: "text", height: content }, { elementId: "text", width: content }]));
});


test("nested group fitting is order independent and preserves descendant world corners", () => {
  const normalized = normalizePresentation(groupFixture({ rotation: 25, flipH: true }));
  const outer = normalized.slides[0].elements[0];
  const inner = normalizePresentation(groupFixture({ rotation: 90, flipV: true })).slides[0].elements[0];
  inner.id = "inner"; inner.children[0].id = "leaf";
  outer.children = [inner];
  const rules = ["group", "inner"].map((elementId) => ({ elementId, width: content, height: content }));
  const a = resolveNormalizedLayout(normalized, { intrinsicSizing: rules });
  const b = resolveNormalizedLayout(normalized, { intrinsicSizing: [...rules].reverse() });
  assert.deepEqual(a, b);
  assert.deepEqual(a.analysis.diagnostics, []);
  const fitted = first(a), fittedInner = fitted.children[0];
  for (const [dx, dy] of [[0, 0], [60, 0], [0, 50], [60, 50]]) {
    const before = world(outer, world(inner, { x: inner.children[0].box.x + dx, y: inner.children[0].box.y + dy }));
    const after = world(fitted, world(fittedInner, { x: fittedInner.children[0].box.x + dx, y: fittedInner.children[0].box.y + dy }));
    assert.ok(Math.abs(before.x - after.x) < 1e-8 && Math.abs(before.y - after.y) < 1e-8);
  }
});

test("content fitting preserves geometry relative to the intentionally parent-resized group", () => {
  const normalized = normalizePresentation(groupFixture({ rotation: 25, flipH: true, flipV: true }));
  const reference = structuredClone(normalized);
  reference.slides[0].elements[0].box.height = 200;
  const actual = resolveNormalizedLayout(normalized, { intrinsicSizing: [{ elementId: "group", width: content, height: { type: "parent" } }] });
  const expected = resolveNormalizedLayout(reference, { intrinsicSizing: [{ elementId: "group", width: content }] });
  assert.deepEqual(actual, expected);
});

test("measurement failures and numeric overflow roll back the entire intrinsic batch", () => {
  for (const measure of [() => { throw new Error("unavailable"); }, () => ({ bounds: { x: 0, y: 0, width: NaN, height: 20 } })]) {
    const document = textFixture();
    const result = fit(document, [{ elementId: "text", height: content }], { textMeasurement: { measure } });
    assert.deepEqual(first(result).box, first(resolveLayout(document)).box);
    assert.ok(result.analysis.diagnostics.length > 0);
  }
  const result = fit(textFixture(), [{ elementId: "text", width: { type: "parent", fraction: Number.MAX_VALUE } }]);
  assert.equal(result.analysis.diagnostics[0].code, "intrinsic-size-invalid");
  assert.equal(first(result).box.width, 60);
});

test("group fitting consumes measured table heights without changing cell geometry", () => {
  const normalized = normalizePresentation(groupFixture());
  const document = createPresentation();
  document.addSlide().addElement({ type: "table", id: "table", box: { x: 0, y: 0, width: 50, height: 20 }, columns: [50], rows: [{ height: 20, cells: [{ content: "abcdefghij", margin: 0 }] }] });
  normalized.slides[0].elements[0].children = normalizePresentation(document).slides[0].elements;
  const result = resolveNormalizedLayout(normalized, { textMeasurement: { measure }, tableSizing: [{ elementId: "table", autoRows: [{ row: 0 }] }], intrinsicSizing: [{ elementId: "group", width: content, height: content }] });
  assert.deepEqual(result.analysis.diagnostics, []);
  const group = first(result), table = group.children[0];
  assert.equal(group.box.height, table.box.height);
  assert.equal(group.coordinateSize.height, table.box.height);
  assert.equal(table.cellLayout[0].box.height, table.box.height);
  assert.ok(table.box.height >= 40);
});

test('rotated rectangular text solves physical height before wrapped width and rolls back on overflow', () => {
  for (const textDirection of ['rotate90', 'rotate270']) {
    const deck = createPresentation();
    deck.addSlide().addElement({ id: 'rotated', type: 'text', box: { x: 12, y: 24, width: 200, height: 120 },
      frame: { textDirection, margin: { left: 3, right: 4, top: 5, bottom: 6 } }, content: '日本語 ABC 123' });
    const requests = [];
    const provider = { measure(request) {
      requests.push([request.wrap, request.availableWidth]);
      return { bounds: { x: 0, y: 0, width: request.wrap ? 60 : 100, height: request.wrap ? 40 : 20 } };
    } };
    const run = (maxWidth, maxHeight) => resolveLayout(deck, { textMeasurement: provider,
      intrinsicSizing: [{ elementId: 'rotated', width: { type: 'content', max: maxWidth }, height: { type: 'content', max: maxHeight } }] });
    const result = run(100, 81);
    assert.deepEqual(result.slides[0].elements[0].box, { x: 12, y: 24, width: 47, height: 81 });
    assert.deepEqual(requests.slice(0, 2), [[false, 0], [true, 70]]);
    for (const [maxWidth, maxHeight] of [[46, 81], [100, 70]]) {
      const failed = run(maxWidth, maxHeight);
      assert.deepEqual(failed.slides[0].elements[0].box, { x: 12, y: 24, width: 200, height: 120 });
      assert.ok(failed.analysis.diagnostics.some(d => d.code === 'intrinsic-size-unsatisfiable'));
    }
  }
});

test('East Asian intrinsic sizing uses vertical metrics and rolls back when columns cannot fit', () => {
  const deck = createPresentation();
  deck.addSlide().addElement({ type: 'text', id: 'vertical', box: { x: 12, y: 24, width: 100, height: 120 },
    frame: { textDirection: 'eastAsianVertical', margin: { left: 3, right: 4, top: 5, bottom: 6 } }, content: '中文 ABC' });
  const calls = [];
  const measureEastAsianVertical = request => {
    calls.push([request.wrap, request.availableInlineSize, request.availableBlockSize]);
    return { bounds: { x: 0, y: 0, width: request.wrap ? 40 : 90, height: request.wrap ? 35 : 18 } };
  };
  const run = (max, provider = { measure, measureEastAsianVertical }) => resolveLayout(deck, { textMeasurement: provider,
    intrinsicSizing: [{ elementId: 'vertical', width: { type: 'content', max }, height: { type: 'content', max: 61 } }] });
  assert.deepEqual(first(run(50)).box, { x: 12, y: 24, width: 42, height: 61 });
  assert.deepEqual(calls.slice(0, 2), [[false, 0, 0], [true, 50, 0]]);
  for (const [result, code] of [[run(41), 'intrinsic-size-unsatisfiable'], [run(50, { measure }), 'east-asian-vertical-measurement-unavailable']]) {
    assert.deepEqual(first(result).box, { x: 12, y: 24, width: 100, height: 120 });
    assert.ok(result.analysis.diagnostics.some(d => d.code === code));
  }
});

test('wrapped vertical parallelograms solve constrained insets in every direction', () => {
  for (const textDirection of ['rotate90', 'rotate270', 'eastAsianVertical', 'mongolianVertical', 'stackedVertical', 'stackedVerticalRtl']) {
    for (const axis of ['both', 'width', 'height']) {
      const deck = createPresentation();
      const box = { x: 10, y: 20, width: 100, height: 90 };
      deck.addSlide().addElement({ type: 'shape', shape: 'parallelogram', slant: 0.4, id: 'p', box,
        text: { content: 'ABCD', frame: { textDirection, wrap: true, margin: 4 } } });
      const calls = [];
      const logical = (wrap, inline, block) => {
        calls.push([wrap, inline, block]);
        return { bounds: { x: 0, y: 0, width: !wrap || inline < 45 ? 130 : 45, height: wrap ? 40 : 12 } };
      };
      const run = (width, height) => resolveLayout(deck, {
        intrinsicSizing: [{ elementId: 'p', ...(width && { width }), ...(height && { height }) }],
        textMeasurement: {
          measure: r => logical(r.wrap, r.availableWidth),
          measureVertical: r => logical(r.wrap, r.availableInlineSize, r.availableBlockSize),
        },
      });
      const width = axis === 'height' ? undefined : { type: 'content', max: 150 };
      const height = axis === 'width' ? undefined : { type: 'content', max: 115 };
      const result = run(width, height);
      assert.ok(!result.analysis.diagnostics.some(d => d.code.startsWith('intrinsic-')), `${textDirection} ${axis}`);
      const shape = first(result), inset = resolveShapeTextBox(shape);
      assert.notDeepEqual(shape.box, box);
      if (!width) assert.equal(shape.box.width, box.width);
      if (!height) assert.equal(shape.box.height, box.height);
      assert.ok(inset.width >= 48 - 1e-9 && inset.height >= 53 - 1e-9);
      assert.ok(calls.some(([wrap, inline]) => wrap && inline >= 45 && inline < 130));
      assert.deepEqual(first(run(width && { ...width, max: 40 }, height && { ...height, max: 40 })).box, box);
    }
  }
});

test('Mongolian and stacked sizing preserve logical leading offsets and roll back on cross-axis failure', () => {
  for (const textDirection of ['mongolianVertical', 'stackedVertical', 'stackedVerticalRtl']) {
    const deck = createPresentation();
    const authored = { x: 12, y: 24, width: 100, height: 120 };
    deck.addSlide().addElement({ type: 'text', id: 'v', box: authored, frame: { textDirection, margin: 5 }, content: 'ABCD' });
    const run = max => resolveLayout(deck, { textMeasurement: { measure, measureVertical(request) {
      assert.equal(request.textDirection, textDirection);
      return { bounds: { x: 2, y: 3, width: request.wrap ? 30 : 60, height: request.wrap ? 20 : 10 } };
    } }, intrinsicSizing: [{ elementId: 'v', width: { type: 'content', max }, height: { type: 'content', max: 45 } }] });
    assert.deepEqual(first(run(33)).box, { ...authored, width: 33, height: 45 });
    const failed = run(32);
    assert.deepEqual(first(failed).box, authored);
    assert.ok(failed.analysis.diagnostics.some(d => d.code === 'intrinsic-size-unsatisfiable'));
  }
});

test('unwrapped vertical parallelograms solve both physical axes and verify the inset text rectangle', () => {
  for (const textDirection of ['rotate90', 'rotate270', 'eastAsianVertical', 'mongolianVertical', 'stackedVerticalRtl']) {
    const deck = createPresentation();
    const box = { x: 10, y: 20, width: 160, height: 100 };
    deck.addSlide().addElement({ type: 'shape', shape: 'parallelogram', slant: 0.6, id: 'p', box,
      text: { content: 'ABCD', frame: { textDirection, wrap: false, margin: { left: 4, right: 6, top: 5, bottom: 7 } } } });
    const vertical = request => {
      assert.equal(request.wrap, false);
      return { bounds: { x: 2, y: 3, width: 40, height: 20 } };
    };
    const run = max => resolveLayout(deck, {
      intrinsicSizing: [{ elementId: 'p', width: { type: 'content', max }, height: content }],
      textMeasurement: { measure: vertical, measureVertical: vertical }, diagnostics: { textOverflow: true },
    });
    const solved = run(300), element = first(solved);
    assert.notDeepEqual(element.box, box);
    const inset = resolveShapeTextBox(element);
    assert.ok(inset.width >= 33 - 1e-9);
    assert.ok(inset.height >= 54 - 1e-9);
    assert.ok(!solved.analysis.diagnostics.some(d => d.code === 'text-overflow'));
    const failed = run(20);
    assert.deepEqual(first(failed).box, box);
    assert.ok(failed.analysis.diagnostics.some(d => d.code === 'intrinsic-size-unsatisfiable'));
  }
});

test('unwrapped vertical parallelograms preserve their fixed axis and roll back impossible fits', () => {
  for (const textDirection of ['rotate90', 'rotate270', 'eastAsianVertical', 'mongolianVertical', 'stackedVertical', 'stackedVerticalRtl']) {
    for (const axis of ['width', 'height']) {
      const other = axis === 'width' ? 'height' : 'width';
      for (const [fixed, max, fits] of [[120, 300, true], [120, 140, false], [axis === 'width' ? 50 : 35, 300, false]]) {
        const deck = createPresentation();
        const box = { x: 10, y: 20, width: 120, height: 120, [other]: fixed };
        deck.addSlide().addElement({ type: 'shape', shape: 'parallelogram', slant: 0.6, id: 'p', box,
          text: { content: 'ABCD', frame: { textDirection, wrap: false, margin: { left: 4, right: 6, top: 5, bottom: 7 } } } });
        const measure = () => ({ bounds: { x: 2, y: 3, width: 40, height: 20 } });
        const result = resolveLayout(deck, {
          intrinsicSizing: [{ elementId: 'p', [axis]: { type: 'content', max } }],
          textMeasurement: { measure, measureVertical: measure },
        });
        const element = first(result);
        const message = `${textDirection}, ${axis}, fixed ${fixed}, max ${max}`;
        if (!fits) {
          assert.deepEqual(element.box, box, message);
          assert.ok(result.analysis.diagnostics.some(d => d.code.startsWith('intrinsic-')), message);
          continue;
        }
        assert.deepEqual(element.box, { ...box, [axis]: element.box[axis] }, message);
        assert.ok(element.box[axis] > 120 && element.box[axis] <= max, message);
        const inset = resolveShapeTextBox(element);
        assert.ok(inset.width >= 33 - 1e-9 && inset.height >= 54 - 1e-9, message);
      }
    }
  }
});

test('fixed-width parallelogram height skips the gap between two fitting ranges', () => {
  const deck = createPresentation();
  deck.addSlide().addElement({ type: 'shape', shape: 'parallelogram', slant: 0.8, id: 'p',
    box: { x: 10, y: 20, width: 120, height: 100 },
    text: { content: 'ABC', frame: { textDirection: 'rotate90', wrap: false, margin: 0 } } });
  const measure = () => ({ bounds: { x: 0, y: 0, width: 28, height: 14 } });
  const solve = (min, max) => resolveLayout(deck, {
    intrinsicSizing: [{ elementId: 'p', height: { type: 'content', min, max } }],
    textMeasurement: { measure },
  });
  const early = first(solve(0, 100));
  assert.ok(early.box.height > 0 && early.box.height < 75);
  assert.ok(Math.abs(resolveShapeTextBox(early).height - 28) < 1e-9);
  const later = first(solve(110, 200));
  assert.ok(Math.abs(later.box.height - 168) < 1e-9);
  assert.ok(Math.abs(resolveShapeTextBox(later).height - 28) < 1e-9);
  const impossible = solve(110, 160);
  assert.equal(first(impossible).box.height, 100);
  assert.ok(impossible.analysis.diagnostics.some(d => d.code === 'intrinsic-size-unsatisfiable'));
});

test('wrapped parallelogram rolls back if the candidate fits only during the inline probe', () => {
  const deck = createPresentation();
  const box = { x: 10, y: 20, width: 100, height: 90 };
  deck.addSlide().addElement({ type: 'shape', shape: 'parallelogram', slant: 0.4, id: 'p', box,
    text: { content: 'ABCD', frame: { textDirection: 'eastAsianVertical', wrap: true, margin: 4 } } });
  let candidates = 0;
  const result = resolveLayout(deck, {
    intrinsicSizing: [{ elementId: 'p', width: { type: 'content', max: 150 } }],
    textMeasurement: { measure, measureVertical({ wrap, availableBlockSize }) {
      if (wrap && availableBlockSize > 0) candidates++;
      return { bounds: { x: 0, y: 0, width: wrap ? 45 : 130, height: availableBlockSize > 0 ? 200 : 40 } };
    } },
  });
  assert.deepEqual(first(result).box, box);
  assert.ok(result.analysis.diagnostics.some(d => d.code === 'intrinsic-shape-height-unavailable'));
  assert.ok(candidates > 1, 'at least one proposed inset and the retained frame were measured');
});
