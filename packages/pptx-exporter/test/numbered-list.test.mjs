import assert from "node:assert/strict";
import test from "node:test";
import { createPresentation } from "@pptkit/core";
import { generatePptx } from "../dist/index.js";
import { readGeneratedZip } from "../../../tests/support/read-generated-zip.mjs";

const numbered = value => [{ style: { bullet: { type: "number", style: "alphaUpperPeriod", startAt: value } }, runs: [{ text: "Item" }] }];
const document = (value, owner) => {
  const deck = createPresentation(), slide = deck.addSlide();
  if (owner === "notes") slide.setNotes(numbered(value));
  else if (owner === "shape") slide.addElement({ type: "shape", shape: "rect", box: { x: 0, y: 0, width: 300, height: 100 }, text: { content: numbered(value) } });
  else if (owner === "table") slide.addElement({ type: "table", box: { x: 0, y: 0, width: 300, height: 100 }, columns: [300], rows: [{ height: 100, cells: [{ content: numbered(value) }] }] });
  else slide.addElement({ type: "text", box: { x: 0, y: 0, width: 300, height: 100 }, content: numbered(value) });
  return deck;
};

test("PPTX numbering accepts 1–32767 and rejects out-of-range emitted markers", async () => {
  for (const owner of ["text", "shape", "notes", "table"]) for (const value of [1, 27, 53, 32767]) {
    const parts = readGeneratedZip((await generatePptx(document(value, owner))).bytes);
    const xml = parts.get(owner === "notes" ? "ppt/notesSlides/notesSlide1.xml" : "ppt/slides/slide1.xml").toString();
    assert.ok(xml.includes(`<a:buAutoNum type="alphaUcPeriod" startAt="${value}"/>`));
  }
  for (const owner of ["text", "shape", "notes", "table"]) await assert.rejects(generatePptx(document(32768, owner)), e => e instanceof RangeError && /1 to 32767/.test(e.message));
});
