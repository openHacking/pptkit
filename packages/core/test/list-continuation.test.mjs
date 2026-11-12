import assert from "node:assert/strict";
import test from "node:test";
import { createPresentation, normalizePresentation, resolvePresentationInheritance, validatePresentation } from "../dist/index.js";

const paragraph = (text, startAt, style = "arabicPeriod") => ({ style: { bullet: { type: "number", style, ...(startAt === undefined ? {} : { startAt }) } }, runs: [{ text }] });
function deck(content, owner = "text") {
  const document = createPresentation(), slide = document.addSlide({ id: "s" });
  if (owner === "notes") slide.setNotes(content);
  else if (owner === "table") slide.addElement({ id: "item", type: "table", box: { x: 0, y: 0, width: 300, height: 200 }, columns: [300], rows: [{ height: 200, cells: [{ content }] }] });
  else if (owner === "shape") slide.addElement({ id: "item", type: "shape", shape: "rect", box: { x: 0, y: 0, width: 300, height: 200 }, text: { content } });
  else slide.addElement({ id: "item", type: "text", box: { x: 0, y: 0, width: 300, height: 200 }, content });
  return document;
}
function normalized(document, owner) {
  const slide = normalizePresentation(document).slides[0];
  return owner === "notes" ? slide.notes : owner === "table" ? slide.elements[0].rows[0].cells[0].content : owner === "shape" ? slide.elements[0].text.content : slide.elements[0].content;
}

test("explicit adjacent continuation materializes deterministically for text, shapes, tables and notes", () => {
  for (const owner of ["text", "shape", "table", "notes"]) {
    const input = [paragraph("A", 4), paragraph("", "continue"), paragraph("C", "continue"), paragraph("restart", 20), paragraph("next", "continue"), { runs: [{ text: "plain" }] }, paragraph("new", undefined), paragraph("last", "continue")];
    const original = structuredClone(input), document = deck(input, owner);
    assert.deepEqual(validatePresentation(document), []);
    const result = normalized(document, owner); assert.deepEqual(result.map(p => p.style.bullet.type === "number" ? p.style.bullet.startAt : null), [4, 5, 6, 20, 21, null, 1, 2]);
    assert.deepEqual(input, original); assert.deepEqual(result, normalized(document, owner)); assert.deepEqual(resolvePresentationInheritance(document).presentation, normalizePresentation(document));
  }
});

test("orphan, mismatched, nonexplicit and overflowing continuation fail validation without mutating input", () => {
  const cases = [
    [paragraph("orphan", "continue")],
    [paragraph("first", 3), { runs: [{ text: "gap" }] }, paragraph("orphan", "continue")],
    [paragraph("first", 3, "alphaLowerPeriod"), paragraph("other", "continue", "alphaUpperPeriod")],
    [paragraph("max", Number.MAX_SAFE_INTEGER), paragraph("overflow", "continue")],
  ];
  for (const input of cases) {
    const before = structuredClone(input), document = deck(input);
    assert.ok(validatePresentation(document).some(d => d.code === "invalid-number-continuation"));
    assert.throws(() => normalizePresentation(document), e => e.name === "PresentationValidationError"); assert.deepEqual(input, before);
  }
});

test("explicit restart and deletion change subsequent calculated numbers without hidden state", () => {
  const document = deck([paragraph("a", 7), paragraph("b", "continue"), paragraph("c", "continue")]);
  assert.deepEqual(normalized(document, "text").map(p => p.style.bullet.startAt), [7, 8, 9]);
  const slide = document.slides[0], element = structuredClone(slide.elements[0]); element.content.splice(1, 1); slide.replaceElement(element.id, element);
  assert.deepEqual(normalized(document, "text").map(p => p.style.bullet.startAt), [7, 8]);
  element.content[0].style.bullet.startAt = 100; slide.replaceElement(element.id, element);
  assert.deepEqual(normalized(document, "text").map(p => p.style.bullet.startAt), [100, 101]);
});

test("continuation reports its local declaration and cannot be declared in a preset", () => {
  const document = deck([paragraph("a", 4), paragraph("b", "continue")]);
  const origin = resolvePresentationInheritance(document).origins.find(o => o.target.id === "item" && o.property === "content.1.style.bullet.startAt");
  assert.equal(origin.source.path, "slides.0.elements.0.content.1.style.bullet.startAt");
  const preset = createPresentation({ textStylePresets: { bad: { paragraph: { bullet: { type: "number", startAt: "continue" } } } } });
  assert.ok(validatePresentation(preset).some(d => d.severity === "error"));
});
