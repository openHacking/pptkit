import assert from "node:assert/strict";
import test from "node:test";
import { createPresentation, normalizePresentation, editTable, TableEditError } from "../dist/index.js";

const preserve = { frame: "preserve" }, fit = { frame: "fit-grid" };
const cell = (content, properties = {}) => ({ content, ...properties });
const fixture = () => ({ type: "table", id: "t", name: "table", box: { x: 5, y: 7, width: 120, height: 90 }, columns: [20, 30, 40], rows: [
  { height: 10, cells: [cell("A", { rowSpan: 2, colSpan: 2, style: { fill: { type: "solid", color: "FF0000" }, margin: 5 } }), cell("B")] },
  { height: 20, cells: [cell("C")] }, { height: 30, cells: [cell("D"), cell("E"), cell("F")] },
] });
const full = (rows = 2, columns = 2) => ({ type: "table", box: { x: 0, y: 0, width: columns * 10, height: rows * 10 }, columns: Array(columns).fill(10), rows: Array.from({ length: rows }, (_, r) => ({ height: 10, cells: Array.from({ length: columns }, (_, c) => cell(`${r},${c}`)) })) });
const text = c => typeof c.content === "string" ? c.content : c.content.map(p => p.runs.map(r => r.text).join("")).join("\n");
function grid(table) {
  // Independent physical partition check, not the implementation's anchor walk.
  const result = Array.from({ length: table.rows.length }, () => Array(table.columns.length).fill(null)), anchors = [];
  for (let r = 0; r < table.rows.length; r++) for (let i = 0; i < table.rows[r].cells.length; i++) {
    const c = table.rows[r].cells[i], col = result[r].indexOf(null), rs = c.rowSpan ?? 1, cs = c.colSpan ?? 1;
    assert.ok(col >= 0); assert.ok(r + rs <= table.rows.length && col + cs <= table.columns.length);
    const at = { rowIndex: r, cellIndex: i, columnIndex: col, rowSpan: rs, colSpan: cs, cell: c }; anchors.push(at);
    for (let y = r; y < r + rs; y++) for (let x = col; x < col + cs; x++) { assert.equal(result[y][x], null); result[y][x] = at; }
  }
  assert.ok(result.flat().every(Boolean)); return { result, anchors };
}
const labels = table => grid(table).result.map(row => row.map(at => text(at.cell)));
function frozen(value) { if (!value || typeof value !== "object") return; assert.ok(Object.isFrozen(value)); for (const nested of Object.values(value)) frozen(nested); }
function error(fn, code) { assert.throws(fn, e => { assert.ok(e instanceof TableEditError); assert.equal(e.code, code); assert.ok(e.path.startsWith("/")); return true; }); }

test("row/column insertion expands only merges crossing the insertion point and preserves source content/styles/mapping", () => {
  const source = fixture(), before = structuredClone(source), result = editTable(source, [
    { type: "insert-row", index: 1, height: 11, cell: cell("row") },
    { type: "insert-column", index: 1, width: 13, cell: cell("column") },
  ], fit);
  assert.deepEqual(labels(result.table), [["A", "A", "A", "B"], ["A", "A", "A", "row"], ["A", "A", "A", "C"], ["D", "column", "E", "F"]]);
  assert.deepEqual(result.table.box, { x: 5, y: 7, width: 103, height: 71 });
  assert.deepEqual(result.table.rows[0].cells[0].style, source.rows[0].cells[0].style);
  assert.deepEqual(result.cellMapping.map(m => [m.status, m.target.rowIndex, m.target.columnIndex]), [["retained", 0, 0], ["retained", 0, 3], ["retained", 2, 3], ["retained", 3, 0], ["retained", 3, 2], ["retained", 3, 3]]);
  assert.deepEqual(source, before); frozen(result);
});

test("boundary insertion shifts anchors without swallowing the new row/column and preserve frame stays exact", () => {
  for (const index of [0, 2, 3]) {
    const source = fixture(), row = editTable(source, [{ type: "insert-row", index, height: 9 }], preserve).table;
    assert.deepEqual(row.box, source.box); assert.ok(labels(row)[index].every(value => value === ""));
    const column = editTable(source, [{ type: "insert-column", index, width: 9 }], preserve).table;
    assert.ok(labels(column).every(values => values[index] === ""));
  }
});

test("deleting a merged anchor row/column promotes its content, shrinks spans and reports only fully removed origins deleted", () => {
  const result = editTable(fixture(), [{ type: "delete-row", index: 0 }, { type: "delete-column", index: 0 }], fit);
  assert.deepEqual(labels(result.table), [["A", "C"], ["E", "F"]]);
  assert.equal(result.table.rows[0].cells[0].rowSpan, 1); assert.equal(result.table.rows[0].cells[0].colSpan, 1);
  assert.deepEqual(result.cellMapping.map(m => m.status), ["retained", "deleted", "retained", "deleted", "retained", "retained"]);
  assert.deepEqual(result.cellMapping[0].target, { rowIndex: 0, cellIndex: 0, columnIndex: 0 });
  assert.deepEqual(result.table.box, { x: 5, y: 7, width: 70, height: 50 });
});

test("interior/end deletion preserves merged content and rejects the final row or column atomically", () => {
  const source = fixture(), before = structuredClone(source);
  assert.deepEqual(labels(editTable(source, [{ type: "delete-row", index: 1 }], preserve).table), [["A", "A", "B"], ["D", "E", "F"]]);
  assert.deepEqual(labels(editTable(source, [{ type: "delete-column", index: 2 }], preserve).table), [["A", "A"], ["A", "A"], ["D", "E"]]);
  for (const type of ["delete-row", "delete-column"]) error(() => editTable(full(1, 1), [{ type, index: 0 }], preserve), "invalid-table-operation");
  error(() => editTable(source, [{ type: "insert-row", index: 0, height: 10 }, { type: "delete-column", index: 99 }], fit), "invalid-table-operation");
  assert.deepEqual(source, before);
});

test("guarded row and column deletion protects removed text and preserves surviving merged anchors", () => {
  const source=fixture(),before=structuredClone(source);
  for(const type of ['delete-row','delete-column']) error(()=>editTable(source,[{type,index:2,content:'reject'}],preserve),'table-edit-content-conflict');
  error(()=>editTable(source,[{type:'insert-row',index:0,height:5},{type:'delete-row',index:1,content:'reject'}],fit),'table-edit-content-conflict');
  assert.deepEqual(source,before);
  const safe=fixture();safe.rows[0].cells[1].content='';safe.rows[2].cells[0].content='';
  const result=editTable(safe,[{type:'delete-row',index:0,content:'reject'},{type:'delete-column',index:0,content:'reject'}],fit);
  assert.deepEqual(labels(result.table),[['A','C'],['E','F']]);
  assert.equal(result.cellMapping[0].status,'retained');assert.deepEqual(result.table.rows[0].cells[0].style,safe.rows[0].cells[0].style);
  const defaultResult=editTable(source,[{type:'delete-row',index:2}],preserve),explicit=editTable(source,[{type:'delete-row',index:2,content:'discard'}],preserve);
  assert.deepEqual(explicit,defaultResult);
});

test("guarded deletion rejects empty actionable runs and accepts genuinely empty or implicit cells", () => {
  const source=full(2,2);source.rows[0].cells[0].content=[{runs:[{text:'',action:{type:'url',url:'https://example.com'}}]}];source.rows[0].cells[1].content='';
  error(()=>editTable(source,[{type:'delete-row',index:0,content:'reject'}],preserve),'table-edit-content-conflict');
  source.rows[0].cells=[];
  const result=editTable(source,[{type:'delete-row',index:0,content:'reject'}],preserve);
  assert.deepEqual(labels(result.table),[['1,0','1,1']]);
  error(()=>editTable(source,[{type:'delete-row',index:0,content:'append-paragraphs'}],preserve),'invalid-table-operation');
});

test("reject-data protects explicit formatting, presets and structured empty text", () => {
  for(const type of ['delete-row','delete-column']) for(const payload of [
    {content:'',style:{fill:{type:'solid',color:'FF0000'}}},
    {content:'',style:{}},
    {content:'',textStylePreset:'title'},
    {content:[]},
    {content:[{style:{align:'center'},runs:[{text:''}]}]},
  ]) {
    const source=full(2,2),before=structuredClone(source);source.rows[0].cells[0]=payload;
    if(type==='delete-row') source.rows[0].cells[1].content=''; else source.rows[1].cells[0].content='';
    const snapshot=structuredClone(source);
    error(()=>editTable(source,[{type,index:0,content:'reject-data'}],preserve),'table-edit-content-conflict');
    assert.deepEqual(source,snapshot);
    // The text-only guard deliberately permits styling with no text or action.
    assert.doesNotThrow(()=>editTable(source,[{type,index:0,content:'reject'}],preserve));
    assert.notDeepEqual(before,snapshot);
  }
  const safe=fixture();safe.rows[0].cells[1].content='';safe.rows[2].cells[0].content='';
  const result=editTable(safe,[{type:'delete-row',index:0,content:'reject-data'},{type:'delete-column',index:0,content:'reject-data'}],fit);
  assert.deepEqual(labels(result.table),[['A','C'],['E','F']]);assert.equal(result.cellMapping[0].status,'retained');
  const blanks=full();blanks.rows[0].cells=[];assert.doesNotThrow(()=>editTable(blanks,[{type:'delete-row',index:0,content:'reject-data'}],preserve));
});

test("default merge rejects conflicting content and partial merged-cell selections without changing input", () => {
  const source = full(), before = structuredClone(source), region = { rowIndex: 0, columnIndex: 0, rowCount: 2, columnCount: 2 };
  error(() => editTable(source, [{ type: "merge", region }], preserve), "table-edit-content-conflict");
  error(() => editTable(fixture(), [{ type: "merge", region: { rowIndex: 0, columnIndex: 0, rowCount: 1, columnCount: 2 } }], preserve), "invalid-table-operation");
  const empty = full(); empty.rows[0].cells[1].content = ""; empty.rows[1].cells.forEach(c => c.content = "");
  empty.rows[1].cells[1].style = { margin: 17 };
  const merged = editTable(empty, [{ type: "merge", region }], preserve);
  assert.deepEqual(labels(merged.table), [["0,0", "0,0"], ["0,0", "0,0"]]);
  assert.deepEqual(merged.cellMapping.map(m => m.status), ["retained", "merged", "merged", "merged"]);
  empty.rows[1].cells[1].content = [{ runs: [{ text: "", action: { type: "url", url: "https://example.com" } }] }];
  error(() => editTable(empty, [{ type: "merge", region }], preserve), "table-edit-content-conflict");
  const actionable = editTable(empty, [{ type: "merge", region, content: "append-paragraphs" }], preserve);
  assert.deepEqual(actionable.table.rows[0].cells[0].content.at(-1).runs[0].action, { type: "url", url: "https://example.com" });
  assert.deepEqual(source, before);
});

test("explicit paragraph append preserves row-major rich text, empty paragraphs, actions and compatible presets", () => {
  const source = full(); source.rows[0].cells[0] = cell([{ style: { align: "right" }, runs: [{ text: "A", style: { bold: true }, action: { type: "url", url: "https://example.com" } }] }], { textStylePreset: "same", style: { margin: 4 } });
  source.rows[0].cells[1] = cell("B\r\n\rC", { textStylePreset: "same" });
  source.rows[1].cells[0] = cell([{ runs: [{ text: "", style: { italic: true } }] }], { textStylePreset: "same" });
  source.rows[1].cells[1] = cell("D", { textStylePreset: "same" });
  const operation = { type: "merge", region: { rowIndex: 0, columnIndex: 0, rowCount: 2, columnCount: 2 }, content: "append-paragraphs" };
  const merged = editTable(source, [operation], fit), content = merged.table.rows[0].cells[0].content;
  assert.deepEqual(content.map(p => p.runs.map(r => r.text).join("")), ["A", "B", "", "C", "", "D"]);
  assert.deepEqual(content[0], source.rows[0].cells[0].content[0]); assert.deepEqual(content[4], source.rows[1].cells[0].content[0]);
  assert.deepEqual(merged.table.rows[0].cells[0].style, { margin: 4 });
  const document = createPresentation({ textStylePresets: { same: { run: { fontSize: 14 } } } }); document.addSlide().addElement(merged.table);
  const normalized = normalizePresentation(document).slides[0].elements[0].rows[0].cells[0].content;
  assert.equal(normalized[0].style.align, "right"); assert.equal(normalized[0].runs[0].style.bold, true);
  assert.equal(normalized[0].runs[0].style.fontSize, 14); assert.equal(normalized[4].runs[0].style.italic, true);
  assert.deepEqual(normalized[0].runs[0].action, source.rows[0].cells[0].content[0].runs[0].action);
  source.rows[1].cells[1].textStylePreset = "different";
  error(() => editTable(source, [operation], fit), "table-edit-content-conflict");
});

test("splitting from a covered coordinate preserves content only at the original anchor and clones empty styles into unit cells", () => {
  const source = fixture(), split = editTable(source, [{ type: "split", rowIndex: 1, columnIndex: 1 }], preserve);
  assert.deepEqual(labels(split.table), [["A", "", "B"], ["", "", "C"], ["D", "E", "F"]]);
  assert.equal(grid(split.table).anchors.length, 9);
  assert.deepEqual(split.table.rows[1].cells[0].style, source.rows[0].cells[0].style);
  assert.notEqual(split.table.rows[1].cells[0].style, split.table.rows[1].cells[1].style);
  assert.ok(split.cellMapping.every(m => m.status === "retained"));
  const merged = editTable(full(), [{ type: "merge", region: { rowIndex: 0, columnIndex: 0, rowCount: 2, columnCount: 2 }, content: "append-paragraphs" }, { type: "split", rowIndex: 1, columnIndex: 1 }], preserve);
  assert.deepEqual(merged.cellMapping.map(m => m.status), ["retained", "merged", "merged", "merged"]);
  assert.ok(merged.cellMapping.every(m => m.target.rowIndex === 0 && m.target.columnIndex === 0));
  assert.equal(text(merged.table.rows[0].cells[0]), "0,0\n0,1\n1,0\n1,1");
});

test("implicit blanks are materialized, default row heights freeze against the original frame, and invalid partitions reject", () => {
  const source = full(); delete source.rows[0].height; delete source.rows[1].height; source.rows[0].cells.pop();
  const result = editTable(source, [{ type: "insert-row", index: 1, height: 5 }], fit);
  assert.deepEqual(result.table.rows.map(r => r.height), [10, 5, 10]); assert.equal(result.cellMapping.length, 3); assert.equal(result.table.box.height, 25);
  const malformed = [
    { ...full(), columns: [] }, { ...full(), rows: [] },
    { ...full(), columns: [-1, 10] }, { ...full(), columns: [NaN, 10] },
  ];
  const outside = full(); outside.rows[0].cells[0].rowSpan = 3; malformed.push(outside);
  const crossing = full(2, 3); crossing.rows[0].cells = [cell("X"), cell("M", { rowSpan: 2 })]; crossing.rows[1].cells = [cell("bad", { colSpan: 2 })]; malformed.push(crossing);
  const overflow = full(); overflow.rows[0].cells.push(cell("extra")); malformed.push(overflow);
  const noDimensions = full(); delete noDimensions.box; delete noDimensions.rows[0].height; malformed.push(noDimensions);
  for (const table of malformed) assert.throws(() => editTable(table, [], preserve), TableEditError);
  for (const frame of [preserve, fit]) assert.throws(() => editTable({ ...full(), columns: [Number.MAX_VALUE, Number.MAX_VALUE] }, [], frame), /overflow/);
});

test("descriptors, sparse/custom arrays, cycles, invalid operations and limits fail without invoking accessors", () => {
  let calls = 0; const source = full(); Object.defineProperty(source.rows[0].cells[0], "content", { enumerable: true, get() { calls++; return "getter"; } });
  error(() => editTable(source, [], preserve), "invalid-table-edit-input"); assert.equal(calls, 0);
  const options = { frame: "fit-grid", limits: { get maxCells() { calls++; return 100; } } };
  error(() => editTable(full(), [], options), "invalid-table-edit-input"); assert.equal(calls, 0);
  const sparse = [, { type: "split", rowIndex: 0, columnIndex: 0 }]; sparse.extra = "replaces hole";
  error(() => editTable(full(), sparse, fit), "invalid-table-edit-input");
  const cyclic = full(); cyclic.self = cyclic; error(() => editTable(cyclic, [], fit), "invalid-table-edit-input");
  for (const operations of [
    [{ type: "insert-row", index: 0, height: 0 }], [{ type: "insert-column", index: .5, width: 10 }],
    [{ type: "insert-row", index: 0, height: 1, cell: cell("x", { colSpan: 2 }) }],
    [{ type: "merge", region: { rowIndex: 0, columnIndex: 0, rowCount: 0, columnCount: 1 } }],
    [{ type: "merge", region: { rowIndex: 0, columnIndex: 0, rowCount: 1, columnCount: 1 }, content: "guess" }],
    [{ type: "delete-row", index: 0, width: 1 }], [{ type: "unknown" }],
  ]) assert.throws(() => editTable(full(), operations, fit), TableEditError);
  for (const limits of [{ maxCells: 0 }, { maxDepth: Infinity }, { maxRows: 10001 }, { unknown: 1 }]) assert.throws(() => editTable(full(), [], { ...fit, limits }), TableEditError);
});

test("each transaction budget has an exact success boundary and failure never poisons an independent retry", () => {
  const source = full(), operations = [{ type: "insert-row", index: 1, height: 10 }, { type: "insert-column", index: 1, width: 10 }], expected = editTable(source, operations, fit);
  const caps = { maxRows: 1000, maxColumns: 1000, maxGridSlots: 1000000, maxCells: 100000, maxOperations: 1000, maxDataValues: 1000000, maxTextCharacters: 16000000, maxWorkUnits: 10000000, maxDepth: 128 };
  for (const [name, cap] of Object.entries(caps)) {
    let low = 1, high = cap;
    while (low < high) {
      const middle = Math.floor((low + high) / 2);
      try { editTable(source, operations, { ...fit, limits: { [name]: middle } }); high = middle; }
      catch (error) { assert.equal(error.code, "table-edit-limit-exceeded"); low = middle + 1; }
    }
    assert.deepEqual(editTable(source, operations, { ...fit, limits: { [name]: low } }), expected, name);
    if (low > 1) error(() => editTable(source, operations, { ...fit, limits: { [name]: low - 1 } }), "table-edit-limit-exceeded");
    assert.deepEqual(editTable(source, operations, fit), expected);
  }
});

test("table commits retain document ownership and whole-object event targets while a failed transaction leaves the live root unchanged", () => {
  const document = createPresentation(), slide = document.addSlide(); slide.addElement(fixture());
  slide.setEventTimeline({ stages: [{ id: "stage", target: { pageId: slide.id, elementId: "t" }, trigger: { type: "enter" }, timing: { kind: "span", id: "span", durationMs: 10 } }] });
  const old = slide.elements[0], before = normalizePresentation(document);
  error(() => editTable(old, [{ type: "delete-row", index: 100 }], fit), "invalid-table-operation"); assert.equal(slide.elements[0], old);
  assert.deepEqual(normalizePresentation(document), before);
  const result = editTable(old, [{ type: "insert-row", index: 1, height: 5 }], fit); slide.replaceElement(old.id, result.table);
  assert.equal(slide.elements[0].id, "t"); assert.equal(slide.getEventTimelinePlan().stages[0].target.elementId, "t");
  const copy = document.duplicateSlide(slide.id); document.removeSlide(slide.id);
  assert.notEqual(copy.elements[0].id, "t"); assert.equal(copy.elements[0].rows.length, 4);
});

test("generated mixed partitions obey insertion/deletion inverses and preserve every original anchor in both axes", () => {
  let seed = 0x51957; const random = n => { seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0; return seed % n; };
  for (let sample = 0; sample < 200; sample++) {
    let source = full(2 + random(5), 2 + random(5));
    const r = random(source.rows.length), c = random(source.columns.length), rs = 1 + random(source.rows.length - r), cs = 1 + random(source.columns.length - c);
    source = editTable(source, [{ type: "merge", region: { rowIndex: r, columnIndex: c, rowCount: rs, columnCount: cs }, content: "append-paragraphs" }], preserve).table;
    const baseline = editTable(source, [], preserve).table;
    for (const axis of ["row", "column"]) {
      const length = axis === "row" ? source.rows.length : source.columns.length, index = random(length + 1);
      const operation = axis === "row" ? { type: "insert-row", index, height: 7 } : { type: "insert-column", index, width: 7 };
      const result = editTable(source, [operation, { type: `delete-${axis}`, index }], preserve);
      assert.deepEqual(result.table, baseline); assert.ok(result.cellMapping.every(m => m.status === "retained")); grid(result.table);
    }
  }
});

test("three thousand mixed structural edits agree with an independent physical-slot splice oracle over multi-merge partitions", () => {
  let seed = 0x816533, serial = 0;
  const random = n => { seed ^= seed << 13; seed ^= seed >>> 17; seed ^= seed << 5; return (seed >>> 0) % n; };
  const counts = { "insert-row": 0, "insert-column": 0, "delete-row": 0, "delete-column": 0 };
  const fresh = text => ({ text, id: serial++ });
  function compact(slots, columns, heights) {
    const positions = new Map();
    slots.forEach((row, r) => row.forEach((owner, c) => { const list = positions.get(owner) ?? []; list.push([r, c]); positions.set(owner, list); }));
    const rows = heights.map(height => ({ height, cells: [] }));
    for (const [owner, at] of positions) {
      const y = Math.min(...at.map(v => v[0])), x = Math.min(...at.map(v => v[1])), bottom = Math.max(...at.map(v => v[0])), right = Math.max(...at.map(v => v[1]));
      assert.equal(at.length, (bottom - y + 1) * (right - x + 1));
      rows[y].cells.push({ column: x, content: owner.text, rowSpan: bottom - y + 1, colSpan: right - x + 1 });
    }
    for (const row of rows) row.cells = row.cells.sort((a, b) => a.column - b.column).map(({ column, ...cell }) => cell);
    return { type: "table", box: { x: 0, y: 0, width: 100, height: 100 }, columns: [...columns], rows };
  }
  for (let sample = 0; sample < 100; sample++) {
    let slots = Array.from({ length: 3 + random(4) }, () => Array.from({ length: 3 + sample % 4 }, () => fresh(`Source ${serial}`)));
    for (let merge = 0; merge < 12; merge++) {
      const r = random(slots.length), c = random(slots[0].length), bottom = r + 1 + random(slots.length - r), right = c + 1 + random(slots[0].length - c);
      const selected = new Set(slots.slice(r, bottom).flatMap(row => row.slice(c, right)));
      if (slots.some((row, y) => row.some((owner, x) => selected.has(owner) && (y < r || y >= bottom || x < c || x >= right)))) continue;
      const owner = fresh(`Merged ${serial}`); for (let y = r; y < bottom; y++) for (let x = c; x < right; x++) slots[y][x] = owner;
    }
    const columns = Array(slots[0].length).fill(10), heights = Array(slots.length).fill(10);
    let table = compact(slots, columns, heights);
    for (let step = 0; step < 30; step++) {
      const rowAxis = random(2) === 0, length = rowAxis ? slots.length : slots[0].length, inserting = length === 1 || (length < 10 && random(2) === 0), index = random(length + (inserting ? 1 : 0));
      const origins = grid(table).anchors.map(a => slots[a.rowIndex][a.columnIndex]);
      const operation = inserting ? rowAxis ? { type: "insert-row", index, height: 7 } : { type: "insert-column", index, width: 7 } : { type: rowAxis ? "delete-row" : "delete-column", index };
      counts[operation.type]++;
      if (inserting && rowAxis) {
        const line = slots[0].map((_, c) => index > 0 && index < slots.length && slots[index - 1][c] === slots[index][c] ? slots[index][c] : fresh(""));
        slots.splice(index, 0, line); heights.splice(index, 0, 7);
      } else if (inserting) {
        for (const line of slots) line.splice(index, 0, index > 0 && index < line.length && line[index - 1] === line[index] ? line[index] : fresh(""));
        columns.splice(index, 0, 7);
      } else if (rowAxis) { slots.splice(index, 1); heights.splice(index, 1); }
      else { for (const line of slots) line.splice(index, 1); columns.splice(index, 1); }
      const result = editTable(table, [operation], preserve), expected = compact(slots, columns, heights), live = new Set(slots.flat());
      assert.deepEqual(result.table, expected, `${sample}/${step}`);
      assert.deepEqual(result.cellMapping.map(m => m.status), origins.map(owner => live.has(owner) ? "retained" : "deleted"));
      table = result.table;
    }
  }
  assert.ok(Object.values(counts).every(count => count > 250), JSON.stringify(counts));
});

test("merge rejects conflicting header scopes and split retains the merged header role", () => {
  const source = full();
  source.rows[0].cells[0].headerScope = "column";
  const region = { rowIndex: 0, columnIndex: 0, rowCount: 1, columnCount: 2 };
  const snapshot = JSON.stringify(source);
  error(() => editTable(source, [{ type: "merge", region, content: "append-paragraphs" }], preserve), "table-edit-header-conflict");
  assert.equal(JSON.stringify(source), snapshot);
  source.rows[0].cells[1].headerScope = "column";
  const merged = editTable(source, [{ type: "merge", region, content: "append-paragraphs" }], preserve).table;
  assert.equal(merged.rows[0].cells[0].headerScope, "column");
  const split = editTable(merged, [{ type: "split", rowIndex: 0, columnIndex: 1 }], preserve).table;
  assert.deepEqual(split.rows[0].cells.slice(0, 2).map(cell => cell.headerScope), ["column", "column"]);
});

test("reject-data deletion protects an empty cell's explicit header scope", () => {
  const source = full();
  source.rows[0].cells.forEach(cell => { cell.content = ""; });
  source.rows[0].cells[0].headerScope = "column";
  error(() => editTable(source, [{ type: "delete-row", index: 0, content: "reject-data" }], preserve), "table-edit-content-conflict");
  const result = editTable(source, [{ type: "delete-row", index: 0, content: "discard" }], preserve);
  assert.equal(result.table.rows.length, 1);
});

test("native table style metadata survives structural edits and invalid values reject atomically", () => {
  const source = full();
  source.nativeFirstRowStyle = false;
  source.nativeTableStylePreset = "medium-2-accent-2";
  source.nativeTableStyleFlags = { firstColumn: true, bandedRows: false };
  const result = editTable(source, [{ type: "insert-row", index: 1, height: 10 }], preserve).table;
  assert.equal(result.nativeFirstRowStyle, false);
  assert.equal(result.nativeTableStylePreset, "medium-2-accent-2");
  assert.deepEqual(result.nativeTableStyleFlags, source.nativeTableStyleFlags);
  assert.notEqual(result.nativeTableStyleFlags, source.nativeTableStyleFlags);
  for (const mutation of [
    { nativeFirstRowStyle: "false" },
    { nativeTableStylePreset: "unknown" },
    { nativeTableStyleFlags: { firstColumn: 1 } },
    { nativeTableStyleFlags: { unexpected: true } },
    { nativeTableStyleFlags: null },
  ]) {
    const invalid = { ...source, ...mutation }, before = JSON.stringify(invalid);
    error(() => editTable(invalid, [{ type: "insert-row", index: 1, height: 10 }], preserve), "invalid-table-edit-input");
    assert.equal(JSON.stringify(invalid), before);
  }
});
