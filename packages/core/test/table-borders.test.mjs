import assert from "node:assert/strict";
import test from "node:test";
import { createPresentation, normalizePresentation, validatePresentation, editTable, PresentationValidationError } from "../dist/index.js";
const base = { paint: { type: "solid", color: "112233", opacity: .6 }, width: 3, dash: "dash" };
function fixture(style = {}) {
  const document = createPresentation(), slide = document.addSlide();
  slide.addElement({ id: "table", type: "table", box: { x: 0, y: 0, width: 120, height: 40 }, columns: [60, 60], rows: [{ height: 40, cells: [{ content: "A", style }, { content: "B" }] }] });
  return { document, slide };
}
const table = document => normalizePresentation(document).slides[0].elements[0];

test("sparse direct edge overrides inherit uniform stroke properties while preserving explicit edge presence", () => {
  const { document } = fixture({ stroke: base, borders: { left: { paint: { type: "none" } }, top: { width: 7 }, right: { paint: { type: "solid", color: { theme: "accent2" } }, dash: "dot" } } });
  const style = table(document).rows[0].cells[0].style;
  assert.deepEqual(Object.keys(style.borders), ["left", "right", "top"]);
  assert.deepEqual(style.borders.left.paint, { type: "none" }); assert.equal(style.borders.left.width, 3);
  assert.equal(style.borders.top.width, 7); assert.deepEqual(style.borders.top.paint, base.paint); assert.equal(style.borders.top.dash, "dash");
  assert.deepEqual(style.borders.right.paint, { type: "solid", color: { theme: "accent2" }, opacity: 1 });
  assert.equal(style.borders.bottom, undefined); assert.equal(style.stroke.width, 3);
});

test("ordinary table IR stays unchanged and explicit empty border records remain a deliberate opt-in", () => {
  const ordinary = fixture(), opted = fixture({ borders: {} });
  assert.equal("borders" in table(ordinary.document).rows[0].cells[0].style, false);
  assert.deepEqual(table(opted.document).rows[0].cells[0].style.borders, {});
  const clean = structuredClone(table(opted.document)); delete clean.rows[0].cells[0].style.borders;
  assert.deepEqual(clean, table(ordinary.document));
});

test("invalid edge shapes, widths, paints, dashes, arrows and unknown themes produce validation diagnostics", () => {
  for (const borders of [null, [], { diagonal: {} }, { top: false }, { left: { width: -1 } }, { right: { paint: null } }, { top: { paint: { type: "solid", color: null } } }, { bottom: { paint: { type: "solid", color: { theme: "unknown" } } } }, { left: { dash: "custom" } }, { right: { beginArrow: "arrow" } }, { left: { compound: "double" } }]) {
    const { document } = fixture({ borders });
    assert.ok(validatePresentation(document).some(d => d.severity === "error"));
    assert.throws(() => normalizePresentation(document), PresentationValidationError);
  }
  // Shared paint validation must report malformed paint rather than dereference it.
  for (const paint of [null, 7, {}, { type: "gradient" }, { type: "solid", color: undefined }]) {
    const { document } = fixture({ fill: paint }); assert.ok(validatePresentation(document).some(d => d.code === "invalid-paint"));
  }
});

test("edge data stays detached across authoring, IR, table edits and whole-slide copies", () => {
  const input = { stroke: base, borders: { top: { width: 4 }, bottom: { paint: { type: "none" } } } }, { document, slide } = fixture(input);
  input.borders.top.width = 99;
  const normalized = table(document); assert.equal(normalized.rows[0].cells[0].style.borders.top.width, 4);
  normalized.rows[0].cells[0].style.borders.top.width = 55; assert.equal(table(document).rows[0].cells[0].style.borders.top.width, 4);
  const result = editTable(slide.elements[0], [{ type: "insert-row", index: 1, height: 20 }], { frame: "fit-grid" }); slide.replaceElement("table", result.table);
  const copy = document.duplicateSlide(slide.id); document.removeSlide(slide.id);
  assert.equal(copy.elements[0].rows[0].cells[0].style.borders.top.width, 4);
  assert.ok(Object.isFrozen(copy.elements[0].rows[0].cells[0].style.borders));
});
