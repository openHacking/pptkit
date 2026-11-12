import assert from "node:assert/strict";
import test from "node:test";
import { createPresentation } from "@pptkit/core";
import { resolveLayout } from "../dist/index.js";

const target = { type: "shape", shape: "rect", id: "target", box: { x: 10, y: 20, width: 40, height: 20 }, transform: { rotation: 90, flipH: true } };
function elements() {
  return [target, ...["top", "right", "bottom", "left", "center"].map(anchor => ({ type: "connector", id: anchor, start: { elementId: "target", anchor }, end: { x: 100, y: 100 } }))];
}
function check(items) {
  const expected = { top: [40, 30], right: [30, 10], bottom: [20, 30], left: [30, 50], center: [30, 30] };
  for (const item of items.filter(item => item.type === "connector")) {
    const [x, y] = expected[item.id];
    assert.ok(Math.abs(item.start.x - x) < 1e-10, item.id);
    assert.ok(Math.abs(item.start.y - y) < 1e-10, item.id);
    assert.deepEqual(item.end, { x: 100, y: 100 });
  }
}
test("all local anchors flip before rotation while center and free endpoints stay fixed", () => {
  const p = createPresentation(); const s = p.addSlide(); elements().forEach(e => s.addElement(e));
  check(resolveLayout(p).slides[0].elements);
});
test("group anchors resolve locally without applying the ancestor transform twice", () => {
  const p = createPresentation();
  p.addSlide().addElement({ type: "group", id: "group", box: { x: 200, y: 300, width: 400, height: 200 }, coordinateSize: { width: 200, height: 100 }, transform: { rotation: 35, flipV: true }, children: elements() });
  check(resolveLayout(p).slides[0].elements[0].children);
});
test("inherited target transforms participate in slide connector geometry", () => {
  const p = createPresentation(); p.defineSlideLayout({ id: "template", name: "Template", elements: [target] });
  const s = p.addSlide({ layoutId: "template" }); elements().slice(1).forEach(e => s.addElement(e));
  check(resolveLayout(p).slides[0].elements);
});
