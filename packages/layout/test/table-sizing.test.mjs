import assert from "node:assert/strict";
import test from "node:test";
import { createPresentation, normalizePresentation } from "@pptkit/core";
import { resolveLayout, resolveNormalizedLayout } from "../dist/index.js";

const cell = (height, extra = {}) => ({ content: String(height), style: { margin: 0 }, ...extra });
function fixture(rows = [
  { height: 30, cells: [cell(10), cell(20)] },
  { height: 40, cells: [cell(60), cell(15)] },
  { height: 25, cells: [cell(10), cell(10)] },
]) {
  const document = createPresentation();
  document.addSlide({ id: "slide" }).addElement({ type: "table", id: "table", box: { x: 10, y: 20, width: 100, height: rows.reduce((sum, row) => sum + row.height, 0) }, columns: [50, 50], rows });
  return document;
}
const measure = (request) => ({ bounds: { x: 0, y: 0, width: 10, height: Number(request.paragraphs.map((p) => p.runs.map((r) => r.text).join("")).join("")) } });
const options = (autoRows, extra = {}) => ({ textMeasurement: { measure }, tableSizing: [{ elementId: "table", autoRows, ...extra }] });
const table = (result) => result.slides[0].elements[0];
const heights = (result) => table(result).rows.map(({ height }) => height);

test('rotated merged text sizes the row span and remeasures at the resulting inline width', () => {
  for(const direction of ['rotate90','rotate270']){
    const d=fixture([{height:10,cells:[{content:'R',rowSpan:2,style:{textDirection:direction,margin:{left:5,right:5,top:3,bottom:7}}},cell(10)]},{height:20,cells:[cell(10)]}]);
    const calls=[],provider={measure:r=>{const text=r.paragraphs[0].runs[0].text;calls.push({text,wrap:r.wrap,width:r.availableWidth});return {bounds:text==='R'?{x:0,y:0,width:60,height:10}:{x:0,y:0,width:5,height:10}}}};
    const result=resolveLayout(d,{textMeasurement:provider,tableSizing:[{elementId:'table',autoRows:[{row:0,maxHeight:25},{row:1}]}]});
    assert.deepEqual(heights(result),[25,45]);assert.ok(calls.some(x=>x.text==='R'&&!x.wrap&&x.width===0));assert.ok(calls.some(x=>x.text==='R'&&x.wrap&&x.width===60));
    assert.ok(result.analysis.diagnostics.some(x=>x.code==='table-text-direction-sizing-approximate'));assert.ok(!result.analysis.diagnostics.some(x=>x.code==='table-cell-overflow'));
  }
});

test('rotated row sizing preserves the source on overflow or final measurement failure', () => {
  const make=()=>fixture([{height:20,cells:[{content:'R',style:{textDirection:'rotate90',margin:5}},cell(10)]}]);
  const rule={elementId:'table',autoRows:[{row:0,maxHeight:40}]};
  const provider={measure:r=>({bounds:r.paragraphs[0].runs[0].text==='R'?{x:0,y:0,width:60,height:10}:{x:0,y:0,width:5,height:10}})};
  const result=resolveLayout(make(),{textMeasurement:provider,tableSizing:[rule]});assert.deepEqual(heights(result),[20]);assert.ok(result.analysis.diagnostics.some(x=>x.code==='table-row-overflow'));
  const clipped=resolveLayout(make(),{textMeasurement:provider,tableSizing:[{...rule,overflow:'clip'}]});assert.deepEqual(heights(clipped),[40]);assert.ok(clipped.analysis.diagnostics.some(x=>x.code==='table-cell-overflow'));
  const failure=resolveLayout(make(),{textMeasurement:{measure:r=>{if(r.wrap&&r.paragraphs[0].runs[0].text==='R')throw Error('measurement failed');return provider.measure(r);}},tableSizing:[{elementId:'table',autoRows:[{row:0}]}]});
  assert.deepEqual(heights(failure),[20]);assert.ok(failure.analysis.diagnostics.some(x=>x.code==='text-measurement-failed'));
});

test('rotated text exceeding fixed cross-width cannot silently size a row', () => {
  const d=fixture([{height:20,cells:[{content:'R',style:{textDirection:'rotate270',margin:5}},cell(10)]}]);
  const provider={measure:r=>({bounds:r.paragraphs[0].runs[0].text==='R'?{x:0,y:0,width:10,height:50}:{x:0,y:0,width:5,height:10}})};
  const result=resolveLayout(d,{textMeasurement:provider,tableSizing:[{elementId:'table',autoRows:[{row:0}]}]});assert.deepEqual(heights(result),[20]);assert.ok(result.analysis.diagnostics.some(x=>x.code==='table-cell-overflow'));
});

test('rotated rows can use a finite maximum when a wrapped measurement fits', () => {
  const make=()=>fixture([{height:20,cells:[{content:'R',style:{textDirection:'rotate90',margin:5}},cell(10)]}]);
  const rule={elementId:'table',autoRows:[{row:0,maxHeight:60}]},measure=r=>({bounds:r.paragraphs[0].runs[0].text==='R'?(r.wrap?{x:0,y:0,width:50,height:20}:{x:0,y:0,width:100,height:10}):{x:0,y:0,width:5,height:10}});
  const result=resolveLayout(make(),{textMeasurement:{measure},tableSizing:[rule],diagnostics:{textOverflow:true}});
  assert.deepEqual(heights(result),[60]);assert.ok(result.analysis.diagnostics.some(x=>x.code==='table-text-direction-wrapped-fallback'));assert.ok(!result.analysis.diagnostics.some(x=>x.code==='table-cell-overflow'||x.code==='table-row-overflow'||x.code==='text-overflow'));
  const bad=resolveLayout(make(),{textMeasurement:{measure:r=>{if(r.wrap&&r.paragraphs[0].runs[0].text==='R')throw Error('wrapped measurement unavailable');return measure(r);}},tableSizing:[rule]});
  assert.deepEqual(heights(bad),[20]);assert.ok(bad.analysis.diagnostics.some(x=>x.code==='text-measurement-failed'));
});

test('rotated wrapping fallback never accepts inline or cross-axis overflow', () => {
  for(const [width,height] of [[51,20],[50,41]]){
    const d=fixture([{height:20,cells:[{content:'R',style:{textDirection:'rotate270',margin:5}},cell(10)]}]);
    const measure=r=>({bounds:r.paragraphs[0].runs[0].text==='R'?(r.wrap?{x:0,y:0,width,height}:{x:0,y:0,width:100,height:10}):{x:0,y:0,width:5,height:10}});
    const result=resolveLayout(d,{textMeasurement:{measure},tableSizing:[{elementId:'table',autoRows:[{row:0,maxHeight:60}]}]});
    assert.deepEqual(heights(result),[20]);assert.ok(result.analysis.diagnostics.some(x=>x.code==='table-row-overflow'));assert.ok(!result.analysis.diagnostics.some(x=>x.code==='table-text-direction-wrapped-fallback'));
  }
});

test('rotated wrapping search uses a smaller verified span height without assuming a global optimum', () => {
  const d=fixture([{height:10,cells:[{content:'R',rowSpan:2,style:{textDirection:'rotate90',margin:5}},cell(10)]},{height:10,cells:[cell(10)]}]);
  const measure=r=>{const text=r.paragraphs[0].runs[0].text;return {bounds:text!=='R'?{x:0,y:0,width:5,height:10}:!r.wrap?{x:0,y:0,width:100,height:10}:{x:0,y:0,width:30,height:r.availableWidth>=30?20:50}}};
  const result=resolveLayout(d,{textMeasurement:{measure},tableSizing:[{elementId:'table',autoRows:[{row:0,maxHeight:30},{row:1,maxHeight:30}]}],diagnostics:{textOverflow:true}}),sum=heights(result).reduce((a,b)=>a+b,0);
  assert.ok(sum>=40-1e-9&&sum<40.1,`resolved height ${sum}`);assert.ok(result.analysis.diagnostics.some(x=>x.code==='table-text-direction-wrapped-fallback'));assert.ok(!result.analysis.diagnostics.some(x=>['text-overflow','table-cell-overflow','table-row-overflow'].includes(x.code)));
});

test('compact rotated sizing searches without a finite maximum and retains conservative default', () => {
  const d=fixture([{height:20,cells:[{content:'R',style:{textDirection:'rotate90',margin:5}},cell(10)]}]);
  let calls=0;const measure=r=>{const text=r.paragraphs[0].runs[0].text;if(text==='R')calls++;return {bounds:text!=='R'?{x:0,y:0,width:5,height:10}:!r.wrap?{x:0,y:0,width:100,height:10}:{x:0,y:0,width:30,height:r.availableWidth>=30?20:50}}};
  const rule={elementId:'table',autoRows:[{row:0}]},normal=resolveLayout(d,{textMeasurement:{measure},tableSizing:[rule]});assert.deepEqual(heights(normal),[110]);
  calls=0;const compact=resolveLayout(d,{textMeasurement:{measure},tableSizing:[{...rule,rotatedTextSizing:'compact'}],diagnostics:{textOverflow:true}}),value=heights(compact)[0];
  assert.ok(value>=40-1e-9&&value<40.1,`resolved height ${value}`);assert.ok(calls<=20);assert.ok(compact.analysis.diagnostics.some(x=>x.code==='table-text-direction-wrapped-fallback'));assert.ok(!compact.analysis.diagnostics.some(x=>['text-overflow','table-cell-overflow','table-row-overflow'].includes(x.code)));
  assert.throws(()=>resolveLayout(d,{tableSizing:[{...rule,rotatedTextSizing:'unknown'}]}),TypeError);
});

test("automatic rows fit content while unspecified fixed rows and authored state are preserved", () => {
  const document = fixture();
  const before = normalizePresentation(document);
  const result = resolveLayout(document, options([{ row: 0 }, { row: 1 }]));
  assert.deepEqual(heights(result), [20, 60, 25]);
  assert.deepEqual(table(result).box, { x: 10, y: 20, width: 100, height: 105 });
  assert.deepEqual(result.analysis.diagnostics, []);
  assert.deepEqual(normalizePresentation(document), before);
  assert.deepEqual(result, resolveLayout(document, options([{ row: 0 }, { row: 1 }])));
});

test("column spans and margins determine available width and required row height", () => {
  const document = fixture([{ height: 10, cells: [cell(20, { colSpan: 2, style: { margin: { left: 4, right: 6, top: 3, bottom: 7 } } })] }]);
  const widths = [];
  const result = resolveLayout(document, { ...options([{ row: 0 }]), textMeasurement: { measure: (request) => { widths.push(request.availableWidth); return measure(request); } } });
  assert.deepEqual(heights(result), [30]);
  assert.ok(widths.every((width) => width === 90));
});

test("merged rows grow only eligible rows and redistribute growth after a cap", () => {
  const document = fixture([
    { height: 10, cells: [cell(100, { rowSpan: 3 }), cell(10)] },
    { height: 20, cells: [cell(10)] },
    { height: 10, cells: [cell(10)] },
  ]);
  const result = resolveLayout(document, options([{ row: 0, maxHeight: 25 }, { row: 2 }]));
  assert.deepEqual(heights(result), [25, 20, 55]);
  assert.deepEqual(result.analysis.diagnostics, []);
});

test("overlapping merged constraints retain all lower bounds deterministically", () => {
  const document = fixture([
    { height: 10, cells: [cell(80, { rowSpan: 2 }), cell(10)] },
    { height: 10, cells: [cell(100, { rowSpan: 2 })] },
    { height: 10, cells: [cell(10)] },
  ]);
  const result = resolveLayout(document, options([{ row: 0 }, { row: 1 }, { row: 2 }]));
  const rows = heights(result);
  assert.ok(rows[0] + rows[1] >= 80 - 1e-9);
  assert.ok(rows[1] + rows[2] >= 100 - 1e-9);
  assert.deepEqual(result.analysis.diagnostics, []);
});

test("overflow preserves the entire table by default; clip applies bounded heights with explicit warnings", () => {
  const document = fixture();
  const result = resolveLayout(document, options([{ row: 0 }, { row: 1, maxHeight: 45 }]));
  assert.deepEqual(heights(result), [30, 40, 25]);
  assert.equal(result.analysis.diagnostics[0].code, "table-row-overflow");
  assert.equal(result.analysis.diagnostics[0].path, "slides.0.elements.0.rows.1.cells.0.content");
  const clipped = resolveLayout(document, options([{ row: 0 }, { row: 1, maxHeight: 45 }], { overflow: "clip" }));
  assert.deepEqual(heights(clipped), [20, 45, 25]);
  assert.ok(clipped.analysis.diagnostics.some(({ code }) => code === "table-row-overflow"));
});

test("fixed overflowing rows reject changes to other rows unless clip is selected", () => {
  assert.deepEqual(heights(resolveLayout(fixture(), options([{ row: 0 }]))), [30, 40, 25]);
  assert.deepEqual(heights(resolveLayout(fixture(), options([{ row: 0 }], { overflow: "clip" }))), [20, 40, 25]);
});

test("invalid grids, width mismatch, missing targets and missing rows retain geometry", () => {
  const normalized = normalizePresentation(fixture());
  const bad = structuredClone(normalized);
  bad.slides[0].elements[0].rows[0].cells[0].rowSpan = 4;
  const span = resolveNormalizedLayout(bad, options([{ row: 0 }]));
  assert.ok(span.analysis.diagnostics.some(({ code }) => code === "table-span-invalid"));
  assert.deepEqual(heights(span), [30, 40, 25]);
  bad.slides[0].elements[0].box.width = 101;
  assert.equal(resolveNormalizedLayout(bad, options([{ row: 0 }])).analysis.diagnostics[0].code, "table-geometry-mismatch");
  assert.equal(resolveNormalizedLayout(normalized, options([{ row: 9 }])).analysis.diagnostics[0].code, "table-sizing-row-missing");
  assert.equal(resolveNormalizedLayout(normalized, { tableSizing: [{ elementId: "missing", autoRows: [] }] }).analysis.diagnostics[0].code, "table-sizing-target-missing");
});

test("provider failures never partially resize a table even with clip enabled", () => {
  const document = fixture();
  for (const provider of [undefined, { measure: () => { throw new Error("offline"); } }, { measure: () => ({ bounds: { x: 0, y: 0, width: 1, height: NaN } }) }]) {
    const result = resolveLayout(document, { tableSizing: [{ elementId: "table", autoRows: [{ row: 0 }, { row: 1 }], overflow: "clip" }], ...(provider ? { textMeasurement: provider } : {}) });
    assert.deepEqual(heights(result), [30, 40, 25]);
    assert.ok(result.analysis.diagnostics.length > 0);
  }
});

test("unresolvable horizontal and top overhang produce explicit cell overflow", () => {
  for (const bounds of [{ x: 0, y: 0, width: 100, height: 5 }, { x: 0, y: -2, width: 10, height: 5 }]) {
    const result = resolveLayout(fixture(), { ...options([{ row: 0 }, { row: 1 }]), textMeasurement: { measure: () => ({ bounds }) } });
    assert.deepEqual(heights(result), [30, 40, 25]);
    assert.ok(result.analysis.diagnostics.some(({ code }) => code === "table-cell-overflow"));
  }
});

test("automatic minima apply to empty content and invalid rule configuration rejects", () => {
  const document = fixture([{ height: 10, cells: [cell(0, { colSpan: 2 })] }]);
  assert.deepEqual(heights(resolveLayout(document, options([{ row: 0, minHeight: 8 }]))), [8]);
  for (const rows of [[{ row: -1 }], [{ row: 0 }, { row: 0 }], [{ row: 0, minHeight: 0 }], [{ row: 0, maxHeight: Infinity }], [{ row: 0, minHeight: 10, maxHeight: 5 }]]) assert.throws(() => resolveLayout(document, options(rows)));
  assert.throws(() => resolveLayout(document, { tableSizing: [{ elementId: "table", autoRows: [] }, { elementId: "table", autoRows: [] }] }));
});

test("hidden tables are opt-in and composition consumes the newly measured height", () => {
  const normalized = normalizePresentation(fixture());
  normalized.slides[0].elements[0].hidden = true;
  assert.deepEqual(heights(resolveNormalizedLayout(normalized, options([{ row: 0 }, { row: 1 }]))), [30, 40, 25]);
  assert.deepEqual(heights(resolveNormalizedLayout(normalized, options([{ row: 0 }, { row: 1 }], { includeHidden: true }))), [20, 60, 25]);
  const result = resolveLayout(fixture(), { ...options([{ row: 0 }, { row: 1 }]), composition: [{ type: "align", container: { kind: "slide", id: "slide" }, elementIds: ["table"], vertical: "end" }] });
  assert.equal(table(result).box.y, 435);
});


test("table sizing resolves rich text and removes identical two-phase font diagnostics", () => {
  const document = fixture([{ height: 10, cells: [{ content: [{ runs: [{ text: "Rich", style: { fontSize: 12, bold: true } }, { text: " text", style: { fontSize: 20, italic: true } }] }], colSpan: 2, style: { margin: 0 } }] }]);
  const requests = [];
  const result = resolveLayout(document, { ...options([{ row: 0 }]), textMeasurement: {
    resolveFont: () => ({ family: "Fallback" }),
    measure: (request) => { requests.push(request); return { bounds: { x: 0, y: 0, width: 60, height: 30 } }; },
  } });
  assert.deepEqual(heights(result), [30]);
  assert.equal(requests[0].paragraphs[0].runs[0].style.bold, true);
  assert.equal(requests[0].paragraphs[0].runs[1].style.italic, true);
  assert.equal(result.analysis.diagnostics.filter(({ code }) => code === "font-fallback").length, 2);
});

test("tables inside hidden nested groups retain source paths and explicit local cell geometry", () => {
  const original = normalizePresentation(fixture()).slides[0].elements[0];
  const document = createPresentation();
  document.addSlide({ id: "nested" }).addElement({ type: "group", id: "group", hidden: true, box: { x: 10, y: 10, width: 200, height: 200 }, coordinateSize: { width: 100, height: 100 }, children: [original] });
  const result = resolveLayout(document, options([{ row: 0 }, { row: 1 }], { includeHidden: true }));
  const nested = result.slides[0].elements[0].children[0];
  assert.deepEqual(nested.rows.map(({ height }) => height), [20, 60, 25]);
  assert.deepEqual(nested.cellLayout[2], { rowIndex: 1, cellIndex: 0, columnIndex: 0, box: { x: 0, y: 20, width: 50, height: 60 } });
  const failure = resolveLayout(document, options([{ row: 0 }], { includeHidden: true }));
  assert.match(failure.analysis.diagnostics[0].path, /^slides.0.elements.0.children.0.rows.1/);
});

test("overflowing summed heights and impossible margins retain all authored row geometry", () => {
  const result = resolveLayout(fixture(), { ...options([{ row: 0 }, { row: 1 }, { row: 2 }]), textMeasurement: { measure: () => ({ bounds: { x: 0, y: 0, width: 1, height: Number.MAX_VALUE } }) } });
  assert.deepEqual(heights(result), [30, 40, 25]);
  assert.ok(result.analysis.diagnostics.some(({ code }) => code === "table-sizing-invalid"));
  const document = fixture([{ height: 10, cells: [cell(0, { colSpan: 2, style: { margin: { left: 60, right: 60 } } })] }]);
  const margin = resolveLayout(document, { ...options([{ row: 0 }]), textMeasurement: { measure: () => ({ bounds: { x: 0, y: 0, width: 0, height: 0 } }) } });
  assert.deepEqual(heights(margin), [10]);
  assert.ok(margin.analysis.diagnostics.some(({ code }) => code === "table-cell-overflow"));
});


test("measurement tolerance is explicit and zero tolerance rejects any measured excess", () => {
  const document = fixture([{ height: 10, cells: [cell(5, { colSpan: 2 })] }]);
  const provider = { measure: () => ({ bounds: { x: 0, y: 0, width: 100.015625, height: 5 } }) };
  const allowed = resolveLayout(document, { ...options([{ row: 0 }]), textMeasurement: provider });
  assert.deepEqual(heights(allowed), [5]);
  const strict = resolveLayout(document, { ...options([{ row: 0 }], { tolerance: 0 }), textMeasurement: provider });
  assert.deepEqual(heights(strict), [10]);
  assert.equal(strict.analysis.diagnostics[0].code, "table-cell-overflow");
  for (const tolerance of [-1, NaN, Infinity]) assert.throws(() => resolveLayout(document, options([{ row: 0 }], { tolerance })));
});

test('East Asian merged cells grow their span with vertical metrics and respect row caps', () => {
  const d = fixture([
    { height: 10, cells: [{ content: '中文ABC', rowSpan: 2, style: {
      textDirection: 'eastAsianVertical', margin: { left: 5, right: 5, top: 3, bottom: 7 },
    } }, cell(10)] },
    { height: 20, cells: [cell(10)] },
  ]);
  const calls = [];
  const provider = { measure, measureEastAsianVertical(request) {
    calls.push([request.wrap, request.availableInlineSize, request.availableBlockSize]);
    return { bounds: { x: 0, y: 0, width: 60, height: 20 } };
  } };
  const run = (maxHeight, overflow) => resolveLayout(d, { textMeasurement: provider,
    tableSizing: [{ elementId: 'table', autoRows: [{ row: 0, maxHeight: 25 }, { row: 1, maxHeight }], overflow }] });
  const grown = run(60, 'preserve');
  assert.deepEqual(heights(grown), [25, 45]);
  assert.ok(calls.some(call => JSON.stringify(call) === '[false,0,0]'));
  assert.ok(calls.some(call => JSON.stringify(call) === '[true,60,40]'));
  const failed = run(30, 'preserve');
  assert.deepEqual(heights(failed), [10, 20]);
  assert.ok(failed.analysis.diagnostics.some(d => d.code === 'table-row-overflow'));
  const clipped = run(30, 'clip');
  assert.deepEqual(heights(clipped), [25, 30]);
  assert.ok(clipped.analysis.diagnostics.some(d => d.code === 'table-cell-overflow'));
});
