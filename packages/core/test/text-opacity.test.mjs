import test from "node:test";
import assert from "node:assert/strict";
import { createPresentation, normalizePresentation, validatePresentation } from "../dist/index.js";

const box = { x: 0, y: 0, width: 200, height: 50 };
const content = colorOpacity => [{ runs: [{ text: "Alpha", style: { color: "123456", colorOpacity } }] }];

test("text alpha snapshots, normalization, duplication and notes preserve zero and fractional opacity", () => {
  for (const alpha of [0, .25, .875, 1]) {
    const p = createPresentation(), s = p.addSlide({ notes: content(alpha) });
    const input = content(alpha);
    const element = s.addElement({ type: "text", box, content: input });
    input[0].runs[0].style.colorOpacity = .99;
    assert.equal(element.content[0].runs[0].style.colorOpacity, alpha);
    assert.throws(() => { element.content[0].runs[0].style.colorOpacity = .5; }, TypeError);
    const copy = s.duplicateElement(element.id);
    assert.equal(copy.content[0].runs[0].style.colorOpacity, alpha);
    const normalized = normalizePresentation(p);
    assert.deepEqual(validatePresentation(p), []);
    assert.equal(normalized.slides[0].elements[0].content[0].runs[0].style.colorOpacity, alpha === 1 ? undefined : alpha);
    assert.equal(normalized.slides[0].notes[0].runs[0].style.colorOpacity, alpha === 1 ? undefined : alpha);
    assert.equal(normalized.slides[0].elements[0].opacity, 1);
  }
});

test("alpha inherits through presets, shape text and table cells, and explicit one resets it", () => {
  const p = createPresentation({ textStylePresets: { translucent: { run: { colorOpacity: .25 } } } });
  const s = p.addSlide();
  s.addElement({ type: "text", box, textStylePreset: "translucent", content: [
    { runs: [{ text: "Inherited" }, { text: "Opaque", style: { colorOpacity: 1 } }, { text: "Transparent", style: { colorOpacity: 0 } }] },
  ] });
  s.addElement({ type: "shape", shape: "rect", box, text: { textStylePreset: "translucent", content: "Shape" } });
  s.addElement({ type: "table", box, columns: [200], rows: [{ cells: [{ textStylePreset: "translucent", content: "Cell" }] }] });
  const [text, shape, table] = normalizePresentation(p).slides[0].elements;
  assert.deepEqual(text.content[0].runs.map(r => r.style.colorOpacity), [.25, undefined, 0]);
  assert.equal(shape.text.content[0].runs[0].style.colorOpacity, .25);
  assert.equal(table.rows[0].cells[0].content[0].runs[0].style.colorOpacity, .25);
  assert.deepEqual(validatePresentation(p), []);
});

test("invalid alpha is diagnosed for direct runs, presets, shape text, cells and notes", () => {
  for (const alpha of [-1, 1.01, NaN, Infinity, "0.5", null]) {
    const p = createPresentation({ textStylePresets: { invalid: { run: { colorOpacity: alpha } } } });
    const s = p.addSlide({ notes: content(alpha) });
    s.addElement({ type: "text", box, content: content(alpha) });
    s.addElement({ type: "shape", shape: "rect", box, text: { content: content(alpha) } });
    s.addElement({ type: "table", box, columns: [200], rows: [{ cells: [{ content: content(alpha) }] }] });
    const issues = validatePresentation(p).filter(d => d.path.endsWith(".colorOpacity"));
    assert.equal(issues.length, 5, String(alpha));
    assert.ok(issues.every(d => d.code === "invalid-opacity"));
  }
});

test("opaque defaults preserve the existing normalized IR shape", () => {
  const p = createPresentation(), s = p.addSlide();
  s.addElement({ type: "text", box, content: "Default" });
  s.addElement({ type: "text", box, content: content(1) });
  const normalized = normalizePresentation(p);
  assert.ok(normalized.slides[0].elements.every(e => !Object.hasOwn(e.content[0].runs[0].style, "colorOpacity")));
});
