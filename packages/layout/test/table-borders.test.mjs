import assert from "node:assert/strict";
import test from "node:test";
import { createPresentation, normalizePresentation } from "@pptkit/core";
import { resolveLayout, resolveNormalizedLayout, resolveTableBorders, TableBorderError } from "../dist/index.js";
const line = (color = "FF0000", width = 1, dash = "solid", opacity = 1) => ({ paint: { type: "solid", color, opacity }, width, dash });
function fixture(cells = [{ content: "A", style: { borders: {} } }, { content: "B" }], rows) {
  const document = createPresentation(), slide = document.addSlide();
  slide.addElement({ id: "table", type: "table", box: { x: 10, y: 20, width: 100, height: 60 }, columns: [40, 60], rows: rows ?? [{ height: 60, cells }] });
  return document;
}
const source = document => normalizePresentation(document).slides[0].elements[0];
const resolved = document => resolveLayout(document).slides[0].elements[0];
const boundary = document => resolved(document).borderGrid.vertical[1];

test("shared edges choose visible, explicit, width, dash then top/left in a deterministic order", () => {
  const pair = (a, b) => fixture([{ content: "A", style: a }, { content: "B", style: b }]);
  assert.equal(boundary(pair({ stroke: line("AA0000", 4), borders: { right: { paint: { type: "none" } } } }, { stroke: line("00BB00", 1) })).paint.color, "00BB00");
  assert.equal(boundary(pair({ stroke: line("AA0000", 9), borders: {} }, { stroke: line("0000BB", 1), borders: { left: line("00BB00", 1) } })).paint.color, "00BB00");
  assert.equal(boundary(pair({ borders: { right: line("AA0000", 2) } }, { borders: { left: line("00BB00", 3) } })).paint.color, "00BB00");
  assert.equal(boundary(pair({ borders: { right: line("AA0000", 3, "dot") } }, { borders: { left: line("00BB00", 3, "solid") } })).paint.color, "00BB00");
  assert.equal(boundary(pair({ borders: { right: line("AA0000", 3) } }, { borders: { left: line("00BB00", 3) } })).paint.color, "AA0000");
  assert.equal(boundary(pair({ borders: { right: line("AA0000", 99, "solid", 0) } }, { borders: { left: line("00BB00", 1) } })).paint.color, "00BB00");
  const horizontal = fixture(undefined, [
    { height: 30, cells: [{ content: "Upper", colSpan: 2, style: { borders: { bottom: line("AA0000", 3) } } }] },
    { height: 30, cells: [{ content: "Lower", colSpan: 2, style: { borders: { top: line("00BB00", 3) } } }] },
  ]);
  assert.equal(resolved(horizontal).borderGrid.horizontal[2].paint.color, "AA0000");
  assert.equal(resolved(horizontal).borderGrid.horizontal[3].paint.color, "AA0000");
});

test("merge interiors are invisible and exterior sides split at adjacent grid boundaries with exact geometry", () => {
  const rows = [{ height: 20, cells: [{ content: "Merged", rowSpan: 2, style: { stroke: line("FF0000", 1), borders: {} } }, { content: "B", style: { borders: { left: line("00FF00", 2) } } }] }, { height: 40, cells: [{ content: "C", style: { borders: { left: line("0000FF", 3) } } }] }];
  const table = resolved(fixture(undefined, rows)), grid = table.borderGrid;
  assert.equal(grid.horizontal[2].paint.type, "none"); assert.equal(grid.horizontal[2].width, 0);
  assert.equal(grid.vertical[1].paint.color, "00FF00"); assert.equal(grid.vertical[4].paint.color, "0000FF");
  assert.deepEqual(table.cellLayout[0], { rowIndex: 0, cellIndex: 0, columnIndex: 0, box: { x: 0, y: 0, width: 40, height: 60 } });
  assert.deepEqual(table.cellLayout[2].box, { x: 40, y: 20, width: 60, height: 40 });
});

test("plain tables retain legacy output; direct helper is explicit, detached and agrees with automatic opt-in", () => {
  const ordinary = fixture([{ content: "A" }, { content: "B" }]); assert.equal(resolved(ordinary).borderGrid, undefined); assert.equal(resolved(ordinary).cellLayout, undefined);
  const document = fixture(), before = normalizePresentation(document), table = resolved(document), direct = resolveTableBorders(source(document));
  assert.deepEqual(table.borderGrid, direct.grid); assert.deepEqual(table.cellLayout, direct.cells);
  direct.grid.vertical[1].width = 100; assert.notEqual(table.borderGrid.vertical[1].width, 100);
  assert.deepEqual(normalizePresentation(document), before); assert.deepEqual(resolveNormalizedLayout(JSON.parse(JSON.stringify(before))), resolveLayout(document));
});

test("uniform dashed strokes resolve physical borders without explicit per-edge overrides", () => {
  for (const dash of ["dash", "dot", "dashDot"]) {
    const document = fixture([{ content: "A", style: { stroke: line("AABBCC", 2, dash, 0.5) } }, { content: "B", style: { stroke: line("123456", 1) } }]);
    const before = normalizePresentation(document), table = resolved(document), direct = resolveTableBorders(source(document));
    assert.deepEqual(table.borderGrid, direct.grid);
    assert.deepEqual(table.cellLayout, direct.cells);
    assert.equal(table.borderGrid.horizontal[0].dash, dash);
    assert.equal(table.borderGrid.horizontal[0].paint.opacity, 0.5);
    assert.equal(table.borderGrid.vertical[1].dash, dash);
    assert.deepEqual(normalizePresentation(document), before);
  }
});

test("grid, edge and work quotas have exact boundaries and share accounting across all opt-in tables", () => {
  const table = source(fixture()), expected = resolveTableBorders(table), caps = { maxGridSlots: 100000, maxEdges: 250000, maxWorkUnits: 2000000 };
  for (const [key, cap] of Object.entries(caps)) {
    let lo = 1, hi = cap;
    while (lo < hi) { const mid = Math.floor((lo + hi) / 2); try { resolveTableBorders(table, { [key]: mid }); hi = mid; } catch (e) { assert.equal(e.code, "table-border-limit-exceeded"); lo = mid + 1; } }
    assert.deepEqual(resolveTableBorders(table, { [key]: lo }), expected);
    if (lo > 1) assert.throws(() => resolveTableBorders(table, { [key]: lo - 1 }), e => e instanceof TableBorderError && e.code === "table-border-limit-exceeded");
  }
  const document = fixture(); document.addSlide().addElement({ ...document.slides[0].elements[0], id: "other" });
  assert.throws(() => resolveLayout(document, { tableBorderLimits: { maxGridSlots: 2 } }), e => e.code === "table-border-limit-exceeded" && e.elementId === "other");
  assert.equal(resolveLayout(document).slides.length, 2);
});

test("raw normalized invalid spans, dimensions, style and limit records reject without partial geometry", () => {
  const baseline = source(fixture());
  for (const mutate of [t => t.rows[0].cells[0].rowSpan = 2, t => t.rows[0].cells[0].colSpan = 3, t => t.columns[0] = 0, t => t.rows[0].height = Infinity, t => t.rows[0].cells[0].style.borders = { unknown: {} }, t => t.rows[0].cells[0].style.stroke.width = -1, t => t.rows[0].cells[0].style.stroke.paint = null, t => t.rows[0].cells[0].style.stroke.dash = "unknown"]) {
    const invalid = structuredClone(baseline); mutate(invalid); assert.throws(() => resolveTableBorders(invalid), TableBorderError);
  }
  let calls = 0;
  for (const limits of [{ maxGridSlots: 0 }, { maxEdges: Infinity }, { unknown: 1 }, { get maxEdges() { calls++; return 100; } }]) assert.throws(() => resolveTableBorders(baseline, limits));
  assert.equal(calls, 0); assert.deepEqual(resolveTableBorders(baseline), resolveTableBorders(baseline));
});

test("missing cells stay unstyled, nested groups preserve local edge geometry and inherited layout tables resolve", () => {
  const document = fixture([{ content: "A", style: { borders: { right: line("FF0000", 2) } } }]);
  const table = resolved(document); assert.equal(table.borderGrid.vertical[1].paint.color, "FF0000"); assert.equal(table.borderGrid.vertical[2].paint.type, "none");
  const nested = createPresentation(); const original = document.slides[0].elements[0];
  nested.addSlide().addElement({ id: "group", type: "group", box: { x: 0, y: 0, width: 200, height: 120 }, coordinateSize: { width: 100, height: 60 }, transform: { rotation: 30 }, children: [original] });
  assert.deepEqual(resolveLayout(nested).slides[0].elements[0].children[0].borderGrid, table.borderGrid);
  assert.deepEqual(resolveLayout(nested).slides[0].elements[0].children[0].cellLayout, table.cellLayout);
});

test("frame/grid disagreement produces a located fidelity warning without silently changing authored dimensions", () => {
  const document = fixture(), slide = document.slides[0], table = structuredClone(slide.elements[0]); table.box.width = 120; slide.replaceElement(table.id, table);
  const result = resolveLayout(document), warning = result.analysis.diagnostics.find(d => d.code === "table-border-frame-grid-mismatch");
  assert.equal(warning.elementId, "table"); assert.equal(warning.slideId, slide.id); assert.equal(warning.bounds.width, 100); assert.equal(warning.available.width, 120);
  assert.equal(result.slides[0].elements[0].box.width, 120); assert.equal(slide.elements[0].box.width, 120);
});
