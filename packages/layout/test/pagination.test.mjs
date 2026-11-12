import assert from "node:assert/strict";
import test from "node:test";
import { createPresentation, normalizePresentation } from "@pptkit/core";
import { resolveLayout, resolveNormalizedLayout } from "../dist/index.js";
import { cellGrid } from "../dist/table/cell-grid.js";
import { planTablePages } from "../dist/table/paginate-tables.js";

const area = { x: 20, y: 30, width: 100, height: 55 };
const rule = (extra = {}) => ({ elementId: "table", area, headerRows: 1, ...extra });
function fixture(count = 7) {
  const document = createPresentation({ size: { width: 300, height: 200 } });
  const slide = document.addSlide({ id: "source", notes: "Source notes", section: "Data", tags: ["report"], customData: { owner: "test" } });
  slide.addElement({ type: "shape", id: "decoration", shape: "rect", box: { x: 0, y: 0, width: 5, height: 5 } });
  slide.addElement({ type: "table", id: "table", box: { x: 10, y: 20, width: 100, height: 10 + (count - 1) * 20 }, columns: [50, 50], rows: Array.from({ length: count }, (_, i) => ({ height: i === 0 ? 10 : 20, cells: [{ content: `L${i}` }, { content: `R${i}` }] })) });
  document.addSlide({ id: "after" }).addElement({ type: "shape", id: "back", shape: "rect", action: { type: "slide", slideId: "source" }, box: { x: 0, y: 0, width: 10, height: 10 } });
  return document;
}
const table = (layout, index = 0) => layout.slides[index].elements.find((element) => element.type === "table");

test("pagination repeats headers, preserves every body row once and inserts continuations in order", () => {
  const document = fixture(), before = normalizePresentation(document);
  const result = resolveLayout(document, { tablePagination: [rule()] });
  assert.deepEqual(result.analysis.diagnostics, []);
  assert.equal(result.slideCount, 4);
  assert.equal(result.slides[0].id, "source"); assert.equal(result.slides.at(-1).id, "after");
  assert.deepEqual(result.pagination.fragments.map((fragment) => fragment.sourceRowIndices), [[0, 1, 2], [0, 3, 4], [0, 5, 6]]);
  assert.deepEqual(result.pagination.fragments.map((fragment) => fragment.repeatedHeaderRowCount), [0, 1, 1]);
  for (let i = 0; i < 3; i++) {
    assert.deepEqual(table(result, i).box, { x: 20, y: 30, width: 100, height: 50 });
    assert.equal(table(result, i).cellLayout.length, 6);
    assert.deepEqual(cellGrid(table(result, i)).invalid, []);
  }
  assert.deepEqual(normalizePresentation(document), before);
  assert.deepEqual(result, resolveLayout(document, { tablePagination: [rule()] }));
  assert.equal(resolveLayout(document).pagination, undefined);
});

test("continuations inherit delivery metadata and layout but do not duplicate notes or source siblings", () => {
  const result = resolveLayout(fixture(), { tablePagination: [rule()] });
  assert.equal(result.slides[0].elements.length, 2);
  for (const slide of result.slides.slice(1, 3)) {
    assert.equal(slide.elements.length, 1); assert.deepEqual(slide.notes, []);
    assert.equal(slide.section, "Data"); assert.deepEqual(slide.tags, ["report"]);
    assert.deepEqual(slide.customData, { owner: "test" }); assert.equal(slide.layoutId, result.slides[0].layoutId);
  }
  assert.ok(result.slides[0].notes.length);
  assert.equal(result.slides.at(-1).elements[0].action.slideId, "source");
});

test("rowspan blocks are indivisible and header boundary crossings are diagnosed", () => {
  const normalized = normalizePresentation(fixture());
  const source = table(normalized);
  source.rows[1].cells[0].rowSpan = 2; source.rows[2].cells.shift();
  const result = resolveNormalizedLayout(normalized, { tablePagination: [rule()] });
  assert.deepEqual(result.pagination.fragments[0].sourceRowIndices, [0, 1, 2]);
  assert.equal(table(result).rows[1].cells[0].rowSpan, 2);
  assert.deepEqual(cellGrid(table(result)).invalid, []);
  const failed = resolveNormalizedLayout(normalized, { tablePagination: [rule({ headerRows: 2 })] });
  assert.equal(failed.analysis.diagnostics[0].code, "pagination-header-merged");
  assert.deepEqual(failed.slides, resolveNormalizedLayout(normalized).slides);
});

test("repeated column headers and merged row headers retain their semantics on each fragment", () => {
  const normalized = normalizePresentation(fixture());
  const source = table(normalized);
  source.rows[0].cells.forEach((cell) => { cell.headerScope = "column"; });
  source.rows[3].cells[0].headerScope = "row";
  source.rows[3].cells[0].rowSpan = 2;
  source.rows[4].cells.shift();
  const before = structuredClone(normalized);
  const result = resolveNormalizedLayout(normalized, { tablePagination: [rule()] });
  assert.deepEqual(result.analysis.diagnostics.map(d => d.code), ["pagination-native-row-style-reapplied"]);
  assert.deepEqual(result.pagination.fragments.map((fragment) => fragment.sourceRowIndices), [[0, 1, 2], [0, 3, 4], [0, 5, 6]]);
  for (let i = 0; i < 3; i++) {
    const fragment = table(result, i);
    assert.deepEqual(fragment.rows[0].cells.map((cell) => cell.headerScope), ["column", "column"]);
    assert.deepEqual(cellGrid(fragment).invalid, []);
  }
  const continuation = table(result, 1);
  assert.equal(continuation.rows[1].cells[0].headerScope, "row");
  assert.equal(continuation.rows[1].cells[0].rowSpan, 2);
  assert.equal(continuation.rows[2].cells.length, 1);
  assert.deepEqual(normalized, before);
  continuation.rows[0].cells[0].headerScope = "row";
  assert.equal(table(result, 0).rows[0].cells[0].headerScope, "column");
  assert.equal(table(result, 2).rows[0].cells[0].headerScope, "column");
});

test("explicit native row-style switches warn only after a successful multi-page split", () => {
  const normalized = normalizePresentation(fixture()), source = table(normalized);
  source.nativeTableStyleFlags = { lastRow: true, bandedRows: true };
  const before = structuredClone(normalized), result = resolveNormalizedLayout(normalized, { tablePagination: [rule()] });
  assert.deepEqual(result.analysis.diagnostics.map(d => [d.code, d.slideId, d.elementId]), [["pagination-native-row-style-reapplied", "source", "table"]]);
  for (let index = 0; index < 3; index++) assert.deepEqual(table(result, index).nativeTableStyleFlags, source.nativeTableStyleFlags);
  assert.deepEqual(normalized, before);
  const failed = resolveNormalizedLayout(normalized, { tablePagination: [rule({ maxPages: 1 })] });
  assert.deepEqual(failed.analysis.diagnostics.map(d => d.code), ["pagination-page-limit"]);
  const single = resolveNormalizedLayout(normalized, { tablePagination: [rule({ area: { ...area, height: 200 } })] });
  assert.deepEqual(single.analysis.diagnostics, []);
  source.nativeFirstRowStyle = false;
  source.nativeTableStyleFlags = { lastRow: false, bandedRows: false };
  assert.deepEqual(resolveNormalizedLayout(normalized, { tablePagination: [rule()] }).analysis.diagnostics, []);
});

test("implicit native defaults warn when first-row or band positions can change", () => {
  const normalized = normalizePresentation(fixture());
  const codes = extra => resolveNormalizedLayout(normalized, { tablePagination: [rule(extra)] }).analysis.diagnostics.map(d => d.code);
  assert.deepEqual(codes({}), [], "two body rows per page preserve band parity");
  assert.deepEqual(codes({ area: { ...area, height: 30 } }), ["pagination-native-row-style-reapplied"]);
  assert.deepEqual(codes({ headerRows: 0 }), ["pagination-native-row-style-reapplied"]);
  assert.deepEqual(codes({ area: { ...area, height: 200 } }), [], "one page has no rebasing");
  assert.deepEqual(codes({ area: { ...area, height: 30 }, maxPages: 1 }), ["pagination-page-limit"]);
  table(normalized).rows[1].cells[0].rowSpan = 2; table(normalized).rows[2].cells.shift();
  assert.deepEqual(codes({}), ["pagination-native-row-style-reapplied"], "merged rows have uncertain native band placement");
});

test("source-boundary policy retains first and last native row regions on their source rows", () => {
  const normalized = normalizePresentation(fixture()), source = table(normalized);
  source.nativeTableStyleFlags = { lastRow: true, bandedRows: false, firstColumn: true };
  const before = structuredClone(normalized);
  const withoutHeaders = resolveNormalizedLayout(normalized, { tablePagination: [rule({ headerRows: 0, nativeRowRegions: "source-boundaries" })] });
  assert.deepEqual(withoutHeaders.analysis.diagnostics, []);
  const parts = withoutHeaders.pagination.fragments.map(f => withoutHeaders.slides.find(s => s.id === f.slideId).elements.find(e => e.type === "table"));
  assert.equal(parts[0].nativeFirstRowStyle, undefined);
  assert.ok(parts.slice(1).every(part => part.nativeFirstRowStyle === false));
  assert.deepEqual(parts.map(part => part.nativeTableStyleFlags.lastRow), [false, false, true]);
  assert.ok(parts.every(part => part.nativeTableStyleFlags.firstColumn));
  const headers = resolveNormalizedLayout(normalized, { tablePagination: [rule({ nativeRowRegions: "source-boundaries" })] });
  assert.ok(headers.pagination.fragments.every(f => headers.slides.find(s => s.id === f.slideId).elements.find(e => e.type === "table").nativeFirstRowStyle !== false));
  assert.deepEqual(headers.analysis.diagnostics, []);
  const failed = resolveNormalizedLayout(normalized, { tablePagination: [rule({ nativeRowRegions: "source-boundaries", maxPages: 1 })] });
  assert.deepEqual(failed.analysis.diagnostics.map(d => d.code), ["pagination-page-limit"]);
  assert.deepEqual(table(failed).nativeTableStyleFlags, source.nativeTableStyleFlags);
  const single = resolveNormalizedLayout(normalized, { tablePagination: [rule({ nativeRowRegions: "source-boundaries", area: { ...area, height: 200 } })] });
  assert.deepEqual(table(single).nativeTableStyleFlags, source.nativeTableStyleFlags);
  assert.deepEqual(normalized, before);
  source.nativeTableStyleFlags.bandedRows = true;
  const banded = resolveNormalizedLayout(normalized, { tablePagination: [rule({ nativeRowRegions: "source-boundaries" })] });
  assert.deepEqual(banded.analysis.diagnostics.map(d => d.code), ["pagination-native-row-style-reapplied"]);
  assert.match(banded.analysis.diagnostics[0].message, /bands may still restart/);
  assert.deepEqual(before.slides[0].elements.find(e => e.id === "table").nativeTableStyleFlags, { lastRow: true, bandedRows: false, firstColumn: true });
});

test("overlapping row spans create a single transitive block", () => {
  const normalized = normalizePresentation(fixture());
  const source = table(normalized);
  source.rows[1].cells[0].rowSpan = 2;
  source.rows[2].cells.shift(); source.rows[2].cells[0].rowSpan = 2;
  source.rows[3].cells.pop();
  const result = resolveNormalizedLayout(normalized, { tablePagination: [rule({ area: { ...area, height: 75 } })] });
  assert.deepEqual(result.analysis.diagnostics.map(d => d.code), ["pagination-native-row-style-reapplied"]);
  assert.deepEqual(result.pagination.fragments.map((f) => f.sourceRowIndices), [[0, 1, 2, 3], [0, 4, 5, 6]]);
});

test("oversized blocks preserve the whole source by default and explicitly overflow alone", () => {
  const normalized = normalizePresentation(fixture()); table(normalized).rows[3].height = 90;
  const failed = resolveNormalizedLayout(normalized, { tablePagination: [rule()] });
  assert.equal(failed.analysis.diagnostics[0].code, "pagination-row-too-tall");
  assert.deepEqual(failed.pagination.fragments, []); assert.deepEqual(failed.slides, resolveNormalizedLayout(normalized).slides);
  const result = resolveNormalizedLayout(normalized, { tablePagination: [rule({ overflow: "overflow" })] });
  const oversized = result.pagination.fragments.find((f) => f.sourceRowIndices.includes(3));
  assert.deepEqual(oversized.sourceRowIndices, [0, 3]);
  const diagnostic = result.analysis.diagnostics.find((d) => d.code === "pagination-row-overflow");
  assert.equal(diagnostic.slideId, oversized.slideId); assert.equal(diagnostic.elementId, oversized.elementId);
  assert.equal(result.slides.find((slide) => slide.id === oversized.slideId).elements[0].box.height, 100);
});

test("different continuation area and layout are explicit and preserve source identity", () => {
  const normalized = normalizePresentation(fixture());
  const second = structuredClone(normalized.layouts[0]); second.id = "continuation"; second.elements = []; second.background = { type: "solid", color: "#112233", opacity: 1 }; normalized.layouts.push(second);
  const result = resolveNormalizedLayout(normalized, { tablePagination: [rule({ continuation: { layoutId: second.id, area: { x: 50, y: 5, width: 200, height: 95 } } })] });
  assert.equal(result.slideCount, 3);
  assert.deepEqual(result.pagination.fragments[1].sourceRowIndices, [0, 3, 4, 5, 6]);
  assert.equal(result.slides[1].layoutId, second.id); assert.deepEqual(result.slides[1].background, second.background);
  assert.equal(result.slides[1].backgroundSource, "layout"); assert.equal(table(result, 1).box.x, 50); assert.equal(table(result, 1).box.width, 100);
});

test("maxPages and deterministic ID collisions fail before any mutation", () => {
  const document = fixture();
  const failed = resolveLayout(document, { tablePagination: [rule({ maxPages: 2 })] });
  assert.equal(failed.analysis.diagnostics[0].code, "pagination-page-limit"); assert.equal(failed.slideCount, 2);
  const id = resolveLayout(document, { tablePagination: [rule()] }).pagination.fragments[1].elementId;
  document.slides[1].addElement({ type: "shape", id, shape: "rect", box: { x: 0, y: 0, width: 1, height: 1 } });
  const conflict = resolveLayout(document, { tablePagination: [rule()] });
  assert.equal(conflict.analysis.diagnostics[0].code, "pagination-id-conflict");
  assert.deepEqual(conflict.slides, resolveLayout(document).slides);
});

test("ambiguous pagination on one source slide rejects both rules irrespective of order", () => {
  const document = fixture();
  document.slides[0].addElement({ type: "table", id: "other", box: { x: 0, y: 0, width: 100, height: 20 }, columns: [100], rows: [{ height: 20, cells: [{ content: "other" }] }] });
  const rules = [rule(), rule({ elementId: "other", headerRows: 0 })];
  for (const order of [rules, [...rules].reverse()]) {
    const result = resolveLayout(document, { tablePagination: order });
    assert.equal(result.analysis.diagnostics.length, 2);
    assert.ok(result.analysis.diagnostics.every((d) => d.code === "pagination-source-conflict"));
    assert.deepEqual(result.slides, resolveLayout(document).slides);
  }
});

test("hidden source slides and tables require explicit opt-in", () => {
  const normalized = normalizePresentation(fixture()); normalized.slides[0].hidden = true;
  assert.equal(resolveNormalizedLayout(normalized, { tablePagination: [rule()] }).slideCount, 2);
  const result = resolveNormalizedLayout(normalized, { tablePagination: [rule({ includeHidden: true })] });
  assert.equal(result.slideCount, 4); assert.ok(result.slides.slice(0, 3).every((slide) => slide.hidden));
});

test("bad targets, widths, layouts and transforms preserve source tables", () => {
  for (const patch of [{ elementId: "decoration" }, { area: { ...area, width: 90 } }, { headerRows: 20 }, { continuation: { layoutId: "missing" } }]) {
    const result = resolveLayout(fixture(), { tablePagination: [rule(patch)] });
    assert.equal(result.slideCount, 2); assert.ok(result.analysis.diagnostics.length);
  }
  const normalized = normalizePresentation(fixture()); table(normalized).transform.rotation = 90;
  assert.equal(resolveNormalizedLayout(normalized, { tablePagination: [rule()] }).analysis.diagnostics[0].code, "pagination-transform-unsupported");
});

test("header-only and zero-header tables never produce empty continuation pages", () => {
  const headerOnly = resolveLayout(fixture(1), { tablePagination: [rule()] });
  assert.equal(headerOnly.slideCount, 2); assert.deepEqual(headerOnly.pagination.fragments[0].sourceRowIndices, [0]);
  const result = resolveLayout(fixture(), { tablePagination: [rule({ headerRows: 0 })] });
  assert.deepEqual(result.pagination.fragments.flatMap((f) => f.sourceRowIndices), [0, 1, 2, 3, 4, 5, 6]);
  assert.ok(result.pagination.fragments.every((f) => f.sourceRowIndices.length > 0));
});

test("measured row heights feed pagination and each fragment gets local cell geometry", () => {
  const result = resolveLayout(fixture(4), {
    textMeasurement: { measure: () => ({ bounds: { x: 0, y: 0, width: 5, height: 35 } }) },
    tableSizing: [{ elementId: "table", autoRows: [0, 1, 2, 3].map((row) => ({ row })) }],
    tablePagination: [rule({ area: { ...area, height: 100 } })],
  });
  assert.deepEqual(result.analysis.diagnostics.map(d => d.code), ["pagination-native-row-style-reapplied"]);
  assert.equal(result.pagination.fragments.length, 3);
  for (let i = 0; i < 3; i++) {
    assert.equal(table(result, i).rows.length, 2);
    assert.equal(table(result, i).cellLayout[2].box.y, table(result, i).rows[0].height);
  }
});

test("final connector resolution sees first fragment bounds; source row data remains detached", () => {
  const document = fixture();
  document.slides[0].addElement({ type: "connector", id: "edge", start: { elementId: "table", anchor: "bottom" }, end: { x: 200, y: 190 } });
  const result = resolveLayout(document, { tablePagination: [rule()] });
  assert.deepEqual(result.slides[0].elements[2].start, { x: 70, y: 80 });
  table(result, 1).rows[0].cells[0].content[0].runs[0].text = "Changed";
  assert.equal(table(result).rows[0].cells[0].content[0].runs[0].text, "L0");
  assert.equal(table(normalizePresentation(document)).rows[0].cells[0].content[0].runs[0].text, "L0");
});

test("bounded generated tables satisfy row conservation, capacities and closed merges", () => {
  for (let count = 1; count <= 24; count++) for (const capacity of [35, 55, 95]) {
    const source = table(normalizePresentation(fixture(count)));
    const plan = planTablePages(source, rule({ area: { ...area, height: capacity } }));
    assert.ok("pages" in plan);
    assert.deepEqual(plan.pages.flatMap((page, i) => i ? page.rows.slice(1) : page.rows), Array.from({ length: count }, (_, i) => i));
    for (const page of plan.pages) {
      assert.ok(page.height <= capacity);
      const fragment = { ...source, rows: page.rows.map((index) => source.rows[index]) };
      assert.deepEqual(cellGrid(fragment).invalid, []);
    }
  }
});

test("malformed pagination contracts reject before mutation", () => {
  for (const patch of [{ area: { ...area, height: 0 } }, { headerRows: -1 }, { maxPages: 0 }, { overflow: "clip" }, { continuation: [] }, { includeHidden: "yes" }, { nativeRowRegions: "unknown" }]) assert.throws(() => resolveLayout(fixture(), { tablePagination: [rule(patch)] }));
  assert.throws(() => resolveLayout(fixture(), { tablePagination: [rule(), rule()] }));
});

test("independent source slides paginate identically regardless of rule ordering", () => {
  const normalized = normalizePresentation(fixture());
  const next = structuredClone(normalized.slides[0]); next.id = "second"; next.elements = [structuredClone(table(normalized))]; next.elements[0].id = "table2";
  normalized.slides.splice(1, 0, next);
  const a = resolveNormalizedLayout(normalized, { tablePagination: [rule(), rule({ elementId: "table2" })] });
  const b = resolveNormalizedLayout(normalized, { tablePagination: [rule({ elementId: "table2" }), rule()] });
  assert.deepEqual(a.slides, b.slides);
  assert.deepEqual([...a.pagination.fragments].sort((x, y) => x.elementId.localeCompare(y.elementId)), [...b.pagination.fragments].sort((x, y) => x.elementId.localeCompare(y.elementId)));
  assert.equal(a.slideCount, 7);
});

test("invalid merge geometry and nonfinite sums do not produce partial pages", () => {
  for (const alter of [(t) => { t.rows[6].cells[0].rowSpan = 2; }, (t) => { t.rows[1].height = Number.MAX_VALUE; t.rows[1].cells[0].rowSpan = 2; t.rows[2].cells.shift(); t.rows[2].height = Number.MAX_VALUE; }]) {
    const normalized = normalizePresentation(fixture()); alter(table(normalized));
    const result = resolveNormalizedLayout(normalized, { tablePagination: [rule({ overflow: "overflow" })] });
    assert.equal(result.slideCount, 2); assert.deepEqual(result.pagination.fragments, []); assert.ok(result.analysis.diagnostics.length);
  }
});

test("row conservation and closed merges hold for varied merged-block plans", () => {
  for (let span = 1; span <= 5; span++) for (const capacity of [35, 55, 95, 135]) {
    const source = table(normalizePresentation(fixture(13)));
    for (let start = 1; start < source.rows.length; start += span) {
      const length = Math.min(span, source.rows.length - start);
      source.rows[start].cells[0].rowSpan = length;
      for (let row = start + 1; row < start + length; row++) source.rows[row].cells.shift();
    }
    const plan = planTablePages(source, rule({ area: { ...area, height: capacity }, overflow: "overflow" }));
    assert.ok("pages" in plan);
    assert.deepEqual(plan.pages.flatMap((page, i) => i ? page.rows.slice(1) : page.rows), Array.from({ length: 13 }, (_, i) => i));
    for (const page of plan.pages) {
      const fragment = { ...source, rows: page.rows.map((row) => source.rows[row]) };
      assert.deepEqual(cellGrid(fragment).invalid, []);
      assert.equal(page.oversized, page.height > capacity);
    }
  }
});

test("pagination retains authored reading order only on the page that owns the original objects", () => {
  const document = fixture(), slide = document.slides[0];
  slide.setReadingOrder([{ elementIds: ["table", "decoration"] }]);
  const result = resolveLayout(document, { tablePagination: [rule()] });
  assert.equal(result.slideCount, 4);
  assert.deepEqual(result.slides[0].readingOrder, slide.readingOrder);
  assert.ok(result.slides.slice(1).every(page => page.readingOrder === undefined));
  assert.equal(table(result, 0).id, "table");
});
