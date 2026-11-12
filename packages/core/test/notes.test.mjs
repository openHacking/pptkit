import assert from "node:assert/strict";
import test from "node:test";
import { createPresentation, normalizePresentation } from "../dist/index.js";

test("speaker notes replacement snapshots nested input and preserves previous frozen values", () => {
  const p = createPresentation(), s = p.addSlide({ notes: "Original" }), previous = s.notes;
  const input = [{ runs: [{ text: "New", style: { bold: true }, action: { type: "url", url: "https://example.invalid" } }], style: { align: "right" } }];
  s.setNotes(input);
  input[0].runs[0].text = "Mutated"; input[0].runs[0].style.bold = false;
  input[0].runs[0].action.url = "https://changed.invalid"; input[0].style.align = "left";
  assert.equal(previous, "Original");
  assert.equal(s.notes[0].runs[0].text, "New"); assert.equal(s.notes[0].runs[0].style.bold, true);
  assert.equal(s.notes[0].runs[0].action.url, "https://example.invalid"); assert.equal(s.notes[0].style.align, "right");
  assert.throws(() => { s.notes[0].runs[0].text = "Direct"; }, TypeError);
  assert.equal(normalizePresentation(p).slides[0].notes[0].runs[0].text, "New");
});

test("notes can be cleared or explicitly empty without altering slide content or earlier snapshots", () => {
  const p = createPresentation(), s = p.addSlide({ notes: "Text" });
  s.addElement({ type: "text", content: "Slide", box: { x: 0, y: 0, width: 100, height: 40 } });
  const before = normalizePresentation(p), element = s.elements[0];
  s.setNotes("");
  assert.equal(s.notes, ""); assert.equal(normalizePresentation(p).slides[0].notes.length, 1);
  s.setNotes();
  assert.equal(s.notes, undefined); assert.equal(normalizePresentation(p).slides[0].notes[0].runs[0].text, "");
  assert.equal(s.elements[0], element); assert.equal(before.slides[0].notes[0].runs[0].text, "Text");
});

test("newly edited notes participate in copied slide self-link remapping and retain external targets", () => {
  const p = createPresentation(), s = p.addSlide({ id: "source" }); p.addSlide({ id: "other" });
  s.setNotes([{ runs: [{ text: "Self", action: { type: "slide", slideId: s.id } }, { text: "Other", action: { type: "slide", slideId: "other" } }] }]);
  const copy = p.duplicateSlide(s.id);
  assert.equal(copy.notes[0].runs[0].action.slideId, copy.id);
  assert.equal(copy.notes[0].runs[1].action.slideId, "other");
  s.setNotes("Source changed");
  assert.equal(copy.notes[0].runs[0].text, "Self");
});

test("a failed notes snapshot leaves the previous notes intact", () => {
  const p = createPresentation(), s = p.addSlide({ notes: "Before" });
  const input = [{ get runs() { throw new Error("Snapshot failed"); } }];
  assert.throws(() => s.setNotes(input), /Snapshot failed/);
  assert.equal(s.notes, "Before");
});
