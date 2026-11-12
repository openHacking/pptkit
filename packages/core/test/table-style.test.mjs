import assert from "node:assert/strict";
import test from "node:test";
import { createPresentation, normalizePresentation, validatePresentation, PresentationValidationError, editTable } from "../dist/index.js";
const fill = color => ({ type: "solid", color });
const line = (color, width = 1) => ({ paint: fill(color), width });
function fixture(style = {}, rows = 4, columns = 4, extra = {}) {
  const document = createPresentation(extra);
  const table = document.addSlide().addElement({ type: "table", id: "table", box: { x: 0, y: 0, width: columns * 40, height: rows * 20 }, columns: Array(columns).fill(40), rows: Array.from({ length: rows }, (_, r) => ({ height: 20, cells: Array.from({ length: columns }, (_, c) => ({ content: `${r}:${c}` })) })), style });
  return { document, table };
}
const normalized = document => normalizePresentation(document).slides[0].elements[0];
const colors = document => normalized(document).rows.map(row => row.cells.map(c => c.style.fill.color));
const replace = (document, table, mutate) => { const changed = structuredClone(table); mutate(changed); document.slides[0].replaceElement(table.id, changed); return changed; };

test("whole/bands/first-last/corners/local layers resolve propertywise in a fixed order", () => {
  const { document, table } = fixture({
    wholeTable: { fill: fill("111111"), margin: 8, stroke: line("111111", 1), text: { run: { fontSize: 10, bold: true }, paragraph: { align: "center" } } },
    bandedColumns: { first: { fill: fill("222222") }, second: { fill: fill("333333") } },
    bandedRows: { first: { fill: fill("444444") }, second: { fill: fill("555555") } },
    firstColumn: { fill: fill("666666") }, lastColumn: { fill: fill("777777") },
    firstRow: { fill: fill("888888"), margin: { top: 0 } }, lastRow: { fill: fill("999999") },
    topLeft: { fill: fill("AA0000"), text: { run: { bold: false, colorOpacity: 0 } } }, topRight: { fill: fill("BB0000") }, bottomLeft: { fill: fill("CC0000") }, bottomRight: { fill: fill("DD0000") },
  });
  replace(document, table, t => { t.rows[0].cells[0].style = { fill: { type: "none" }, margin: { right: 0 }, borders: { top: { width: 0 } } }; t.rows[0].cells[0].content = [{ runs: [{ text: "Local", style: { fontSize: 12, italic: false } }] }]; });
  const t = normalized(document), c = t.rows[0].cells[0];
  assert.deepEqual(colors(document), [[undefined,"888888","888888","BB0000"],["666666","444444","444444","777777"],["666666","555555","555555","777777"],["CC0000","999999","999999","DD0000"]]);
  assert.deepEqual(c.style.margin, { top: 0, right: 0, bottom: 8, left: 8 });
  assert.equal(c.style.borders.top.width, 0);
  assert.equal(c.content[0].runs[0].style.bold, false); assert.equal(c.content[0].runs[0].style.colorOpacity, 0); assert.equal(c.content[0].runs[0].style.fontSize, 12); assert.equal(c.content[0].style.align, "center");
  assert.equal(t.style, undefined, "authoring rules materialize into canonical cells");
});

test("inner/outer borders follow physical merge boundaries and later uniform/per-edge patches compose", () => {
  const { document, table } = fixture({ wholeTable: { stroke: line("111111", 1) }, borders: { outer: line("FF0000", 3), innerHorizontal: line("00FF00", 2), innerVertical: line("0000FF", 4) }, firstRow: { borders: { bottom: { dash: "dot" } } } }, 3, 3);
  replace(document, table, t => { t.rows[0].cells = [{ content: "Merged", rowSpan: 2, colSpan: 2, style: { stroke: { width: 5 }, borders: { right: { width: 6 } } } }, { content: "Upper" }]; t.rows[1].cells = [{ content: "Lower" }]; });
  const t = normalized(document), c = t.rows[0].cells[0];
  assert.equal(c.style.borders.left.paint.color, "FF0000"); assert.equal(c.style.borders.top.paint.color, "FF0000"); assert.equal(c.style.borders.bottom.paint.color, "00FF00"); assert.equal(c.style.borders.right.paint.color, "0000FF");
  assert.equal(c.style.borders.left.width, 5); assert.equal(c.style.borders.right.width, 6); assert.equal(c.style.borders.bottom.dash, "dot");
});

test("band sizes use physical anchors, exclude configured first/last lines, and degenerate corners have deterministic precedence", () => {
  const { document } = fixture({ bandedRows: { size: 2, first: { fill: fill("AA0000") }, second: { fill: fill("00AA00") } }, firstRow: {}, lastRow: {} }, 8, 1);
  assert.deepEqual(colors(document).flat(), [undefined,"AA0000","AA0000","00AA00","00AA00","AA0000","AA0000",undefined]);
  const single = fixture({ firstRow: { fill: fill("111111") }, lastRow: { fill: fill("222222") }, firstColumn: { fill: fill("333333") }, lastColumn: { fill: fill("444444") }, topLeft: { fill: fill("555555") }, topRight: { fill: fill("666666") }, bottomLeft: { fill: fill("777777") }, bottomRight: { fill: fill("888888") } }, 1, 1);
  assert.equal(colors(single.document)[0][0], "888888");
});

test("implicit blanks participate, merged anchors receive touching edge regions, and text presets remain local overrides", () => {
  const { document, table } = fixture({ wholeTable: { margin: 7, text: { run: { fontFamily: { theme: "heading" }, fontSize: 10, bold: true }, paragraph: { align: "right" } } }, lastRow: { fill: fill("00AA00") }, lastColumn: { fill: fill("0000AA") } }, 3, 3, { textStylePresets: { local: { frame: { margin: { left: 0 } }, run: { fontSize: 14, bold: false }, paragraph: { align: "center" } } } });
  replace(document, table, t => { t.rows[0].cells = [{ content: "A\r\n\rB\n", rowSpan: 3, colSpan: 2, textStylePreset: "local" }]; t.rows[1].cells = []; t.rows[2].cells = []; });
  const t = normalized(document), c = t.rows[0].cells[0];
  assert.deepEqual(t.rows.map(r => r.cells.length), [2,1,1]); assert.equal(c.style.fill.color, "00AA00");
  assert.deepEqual(c.content.map(p => p.runs.map(r => r.text).join("")), ["A","","B",""]);
  assert.equal(c.style.margin.left, 0); assert.equal(c.style.margin.right, 7); assert.equal(c.content[0].runs[0].style.bold, false); assert.equal(c.content[0].runs[0].style.fontSize, 14); assert.deepEqual(c.content[0].runs[0].style.fontFamily, { theme: "heading" }); assert.equal(c.content[0].style.align, "center");
  assert.equal(t.rows[1].cells[0].style.fill.color, "0000AA"); assert.equal(t.rows[2].cells[0].style.fill.color, "00AA00");
  const blank = fixture({ wholeTable: { text: { run: { fontSize: 23, bold: true } } } }, 1, 2);
  replace(blank.document, blank.table, t => { t.rows[0].cells[0].content = []; t.rows[0].cells[1].content = [{ runs: [] }]; });
  for (const c of normalized(blank.document).rows[0].cells) {
    assert.equal(c.content[0].runs[0].text, ""); assert.equal(c.content[0].runs[0].style.fontSize, 23); assert.equal(c.content[0].runs[0].style.bold, true);
  }
});

test("rules are detached, explicit undefined inherits, removing rules restores legacy, edits recompute bands", () => {
  const input = { wholeTable: { fill: fill("111111"), text: { run: { bold: true } } }, bandedRows: { first: { fill: fill("AA0000") }, second: { fill: fill("00AA00") } } };
  const { document, table } = fixture(input, 3, 2), before = normalizePresentation(document);
  input.wholeTable.fill.color = "FFFFFF"; assert.equal(colors(document)[0][0], "AA0000");
  const changed = replace(document, table, t => { t.rows[0].cells[0].style = { fill: undefined }; t.rows[0].cells[0].content = [{ runs: [{ text: "Inherited", style: { bold: undefined } }] }]; });
  assert.equal(normalized(document).rows[0].cells[0].content[0].runs[0].style.bold, true);
  const edited = editTable(changed, [{ type: "insert-row", index: 1, height: 20 }], { frame: "fit-grid" }); document.slides[0].replaceElement(table.id, edited.table);
  assert.deepEqual(colors(document).map(row => row[0]), ["AA0000","00AA00","AA0000","00AA00"]);
  const latest = document.slides[0].elements[0]; replace(document, latest, t => { delete t.style; }); assert.equal(normalized(document).rows[1].cells[0].style.borders, undefined);
  const detached = structuredClone(before); detached.slides[0].elements[0].rows[0].cells[0].style.fill.color = "FFFFFF"; assert.notDeepEqual(normalizePresentation(document), detached);
});

test("closed rule grammar, malformed values and overlapping/out-of-bounds partitions reject through located diagnostics", () => {
  const mutations = [t => t.style = null, t => t.style.unknown = {}, t => t.style.bandedRows = { size: 0 }, t => t.style.bandedColumns = { size: 100001 }, t => t.style.firstRow = { fill: null }, t => t.style.lastRow = { margin: { left: -1 } }, t => t.style.firstColumn = { borders: { diagonal: {} } }, t => t.style.lastColumn = { stroke: { beginArrow: "triangle" } }, t => t.style.topLeft = { text: { run: { bold: 1 } } }, t => t.style.bottomRight = { text: { paragraph: { align: "unknown" } } }, t => t.style.wholeTable = { text: { run: { fontFamily: { theme: "unknown" } } } }, t => t.style.borders = { outer: { width: Infinity } }, t => t.rows[0].cells[0].rowSpan = 5, t => { t.rows[0].cells[1].rowSpan = 2; t.rows[1].cells = [{ content: "overlap", colSpan: 2 }]; }, t => t.rows[0].cells[0].style = { margin: -1 }, t => t.rows[0].cells[0].style = null];
  for (const mutate of mutations) {
    const { document, table } = fixture(); replace(document, table, mutate); const diagnostics = validatePresentation(document);
    assert.ok(diagnostics.some(d => d.code === "invalid-table-style" && d.elementId === "table" && d.path.startsWith("slides.0.elements.0.")), JSON.stringify(diagnostics));
    assert.throws(() => normalizePresentation(document), PresentationValidationError);
  }
});

test("bounded styled grid accepts exact slot boundary and rejects the next slot before normalization", () => {
  const { document, table } = fixture({}, 1, 1);
  replace(document, table, t => { t.columns = Array(100000).fill(1); t.rows[0].cells = []; });
  assert.equal(normalized(document).rows[0].cells.length, 100000);
  const last = document.slides[0].elements[0]; replace(document, last, t => { t.columns.push(1); });
  assert.throws(() => normalizePresentation(document), e => e.diagnostics.some(d => d.code === "invalid-table-style" && /100000/.test(d.message)));
});

test("160 independently partitioned merged grids match geometric region selection, including every edge/corner", () => {
  let seed = 0x15a3f9c7;
  const random = n => { seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0; return seed % n; };
  const names = ["wholeTable","firstColumn","lastColumn","firstRow","lastRow","topLeft","topRight","bottomLeft","bottomRight"];
  for (let round=0;round<160;round++) {
    const rows=1+random(7), columns=1+random(7), rects=[{r:0,c:0,h:rows,w:columns}];
    for (let cut=0;cut<12;cut++) {
      const index=random(rects.length), a=rects[index], horizontal=a.h>1 && (a.w===1 || random(2)===0);
      if (horizontal) { const size=1+random(a.h-1); rects.splice(index,1,{...a,h:size},{...a,r:a.r+size,h:a.h-size}); }
      else if(a.w>1) { const size=1+random(a.w-1); rects.splice(index,1,{...a,w:size},{...a,c:a.c+size,w:a.w-size}); }
    }
    const style={}, codes={};
    for(let i=0;i<names.length;i++) if(random(3)!==0) { codes[names[i]]=(0x100000+i*0x10101).toString(16).toUpperCase(); style[names[i]]={fill:fill(codes[names[i]])}; }
    style.bandedRows={size:1+random(3),first:{fill:fill("AA0000")},second:{fill:fill("00AA00")}};
    style.bandedColumns={size:1+random(3),first:{fill:fill("0000AA")},second:{fill:fill("AAAA00")}};
    const document=createPresentation(), cells=Array.from({length:rows},()=>[]);
    rects.sort((a,b)=>a.r-b.r||a.c-b.c);
    for(const a of rects) cells[a.r].push({content:`${a.r}:${a.c}`,rowSpan:a.h,colSpan:a.w});
    document.addSlide().addElement({type:"table",columns:Array(columns).fill(10),rows:cells.map(cells=>({height:10,cells})),box:{x:0,y:0,width:columns*10,height:rows*10},style});
    const actual=normalized(document);
    for(const a of rects) {
      let expected=codes.wholeTable;
      // Independent geometric rule oracle, without compact-grid traversal.
      for(const [start,total,first,last,band,paints] of [[a.c,columns,"firstColumn","lastColumn",style.bandedColumns,["0000AA","AAAA00"]],[a.r,rows,"firstRow","lastRow",style.bandedRows,["AA0000","00AA00"]]]) {
        if(!(style[first] && start===0) && !(style[last] && start===total-1)) expected=paints[Math.floor((start-(style[first]?1:0))/band.size)%2];
      }
      const touches={firstColumn:a.c===0,lastColumn:a.c+a.w===columns,firstRow:a.r===0,lastRow:a.r+a.h===rows,topLeft:a.c===0&&a.r===0,topRight:a.c+a.w===columns&&a.r===0,bottomLeft:a.c===0&&a.r+a.h===rows,bottomRight:a.c+a.w===columns&&a.r+a.h===rows};
      for(const name of names.slice(1)) if(touches[name]&&style[name]) expected=codes[name];
      const cell=actual.rows[a.r].cells.find(c=>c.content[0].runs[0].text===`${a.r}:${a.c}`);
      assert.equal(cell.style.fill.color,expected,`${round}/${a.r}/${a.c}`);
    }
  }
});

test("native table switches are detached and do not change effective Core regions", () => {
  const { document, table } = fixture({ firstColumn: { fill: fill("123456") }, bandedRows: { first: { fill: fill("FEDCBA") } } });
  const before = normalized(document).rows;
  const flags = { firstColumn: false, lastColumn: true, bandedRows: false };
  replace(document, table, changed => { changed.nativeTableStyleFlags = flags; });
  flags.lastColumn = false;
  const after = normalized(document);
  assert.deepEqual(after.nativeTableStyleFlags, { firstColumn: false, lastColumn: true, bandedRows: false });
  assert.deepEqual(after.rows, before);
  for (const invalid of [null, [], { firstColumn: 1 }, { firstRow: true }, { bandedRows: "off" }]) {
    replace(document, table, changed => { changed.nativeTableStyleFlags = invalid; });
    assert.ok(validatePresentation(document).some(d => d.code === "invalid-native-table-style-flags" && d.path.endsWith(".nativeTableStyleFlags")));
  }
});

test("native preset selection validates known names without changing effective cells", () => {
  const { document, table } = fixture({ firstRow: { fill: fill("123456") } });
  const cells = normalized(document).rows;
  replace(document, table, changed => { changed.nativeTableStylePreset = "medium-2-accent-2"; });
  assert.equal(normalized(document).nativeTableStylePreset, "medium-2-accent-2");
  assert.deepEqual(normalized(document).rows, cells);
  replace(document, table, changed => { changed.nativeTableStylePreset = "unknown"; });
  assert.ok(validatePresentation(document).some(item => item.code === "invalid-native-table-style-preset"));
});
