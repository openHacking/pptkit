import assert from "node:assert/strict";
import test from "node:test";
import { createHash } from "node:crypto";
import { createPresentation, normalizePresentation } from "../packages/core/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { importPptx, compareOpcPackages, inspectOpcClosure } from "../packages/pptx-parser/dist/index.js";
import { createTextOwnerFixture } from "../tests/fixtures/text-owners.mjs";
import { png } from "../packages/pptx-parser/test/image-fixture.mjs";
import { zip } from "../packages/pptx-parser/test/zip-fixture.mjs";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";

const profile = { projection: "content", editSaveProfile: "text-runs-v2" };
const digest = bytes => createHash("sha256").update(bytes).digest("hex");
const repack = parts => zip([...parts].map(([name, body]) => ({ name, body }))).bytes;
async function source() {
  const parts = readGeneratedZip((await generatePptx(createTextOwnerFixture(createPresentation, png))).bytes);
  for (const [part, end] of [["ppt/slides/slide1.xml", "sld"], ["ppt/notesSlides/notesSlide1.xml", "notes"]]) {
    const opaque = '<p:ext uri="urn:text-owner-evidence"><x:opaque xmlns:x="urn:future">Original shape text | Original speaker notes</x:opaque></p:ext>';
    const xml = parts.get(part).toString();
    parts.set(part, Buffer.from(xml.includes("</p:extLst>") ? xml.replace("</p:extLst>", opaque + "</p:extLst>") : xml.replace(`</p:${end}>`, `<p:extLst>${opaque}</p:extLst></p:${end}>`)));
  }
  parts.set("orphan.xml", Buffer.from('<retained xmlns="urn:future">untouched</retained>'));
  return repack(parts);
}
function edit(session, shapeText, notesText) {
  const slide = session.document.slides[0], index = slide.elements.findIndex(e => e.type === "shape" && e.text);
  const shape = structuredClone(slide.elements[index]); shape.text.content[0].runs[0].text = shapeText;
  slide.removeElement(shape.id); slide.insertElement(index, shape);
  const notes = structuredClone(slide.notes); notes[0].runs[0].text = notesText; slide.setNotes(notes);
}
async function save(session) {
  const plan = await session.prepareSave({ mode: "edit" });
  assert.equal(plan.status, "ready", JSON.stringify(plan.report));
  return session.save(plan.plan);
}

test("generated mixed decks retain all chart/workbook/table/image/other-note payloads over three v2 rounds", async () => {
  let input = await source(), previousSize;
  const original = readGeneratedZip(input), graph = await inspectOpcClosure(input, { roots: ["/"] });
  for (let round = 1; round <= 3; round++) {
    const session = await importPptx(input, profile);
    try {
      edit(session, `Shape round ${round} <&🙂`, `Notes round ${round} <&🙂`);
      const result = await save(session), parts = readGeneratedZip(result.bytes);
      assert.equal(result.report.changes.length, 2);
      assert.equal(result.report.changes.find(c => c.source.part.includes("/notesSlides/")).elementId, undefined);
      assert.deepEqual([...parts.keys()].sort(), [...original.keys()].sort());
      for (const [name, payload] of original) if (!["ppt/slides/slide1.xml", "ppt/notesSlides/notesSlide1.xml"].includes(name)) assert.equal(digest(parts.get(name)), digest(payload), name);
      assert.ok(parts.get("ppt/slides/slide1.xml").toString().includes("Original shape text | Original speaker notes"));
      assert.ok(parts.get("ppt/notesSlides/notesSlide1.xml").toString().includes("Original shape text | Original speaker notes"));
      const compared = await compareOpcPackages(input, result.bytes);
      assert.deepEqual(compared.differences.map(d => [d.code, d.part]), [["payload-changed", "/ppt/notesSlides/notesSlide1.xml"], ["payload-changed", "/ppt/slides/slide1.xml"]]);
      const nextGraph = await inspectOpcClosure(result.bytes, { roots: ["/"] });
      assert.deepEqual(nextGraph.relationships, graph.relationships); assert.deepEqual(nextGraph.unreachedParts, graph.unreachedParts);
      const reopened = await importPptx(result.bytes, profile);
      assert.deepEqual(normalizePresentation(reopened.document), normalizePresentation(session.document));
      reopened.dispose();
      if (previousSize !== undefined) assert.equal(result.bytes.length, previousSize);
      previousSize = result.bytes.length; input = result.bytes;
    } finally { session.dispose(); }
  }
});

test("table/notes formatting mutations block a complete v2 plan and exact restoration permits its supported edits", async () => {
  const input = await source(), session = await importPptx(input, profile);
  try {
    edit(session, "Shape supported", "Notes supported");
    const slide = session.document.slides[0], index = slide.elements.findIndex(e => e.type === "table"), table = slide.elements[index];
    const changed = structuredClone(table); changed.rows[0].cells[0].content[0].runs[0].text = "unsupported cell";
    slide.removeElement(table.id); slide.insertElement(index, changed);
    assert.equal((await session.prepareSave({ mode: "edit" })).status, "blocked");
    slide.removeElement(table.id); slide.insertElement(index, table);
    const notes = structuredClone(slide.notes), styled = structuredClone(notes); styled[0].runs[0].style.italic = true; slide.setNotes(styled);
    assert.equal((await session.prepareSave({ mode: "edit" })).status, "blocked");
    slide.setNotes(notes);
    const saved = await save(session); assert.equal(saved.report.changes.length, 2);
    const v1 = await importPptx(saved.bytes, { projection: "content", editSaveProfile: "text-runs-v1" });
    edit(v1, "still outside v1", "still outside v1");
    assert.equal((await v1.prepareSave({ mode: "edit" })).status, "blocked"); v1.dispose();
  } finally { session.dispose(); }
});

test("v2 drawing/notes payload patches preserve unlike UTF-16 byte orders through production candidates", async () => {
  const parts = readGeneratedZip(await source());
  for (const [part, bigEndian] of [["ppt/slides/slide1.xml", false], ["ppt/notesSlides/notesSlide1.xml", true]]) {
    const xml = parts.get(part).toString().replace(/encoding="UTF-8"/i, 'encoding="UTF-16"');
    const encoded = Buffer.from(xml, "utf16le");
    parts.set(part, Buffer.concat([Buffer.from(bigEndian ? [254, 255] : [255, 254]), bigEndian ? encoded.swap16() : encoded]));
  }
  const session = await importPptx(repack(parts), profile);
  try {
    edit(session, "Shape encoded 🙂\r", "Notes encoded 🙂\r");
    const saved = await save(session), output = readGeneratedZip(saved.bytes);
    for (const [part, bigEndian, before, after] of [["ppt/slides/slide1.xml", false, "Original shape text", "Shape encoded 🙂&#xD;"], ["ppt/notesSlides/notesSlide1.xml", true, "Original speaker notes", "Notes encoded 🙂&#xD;"]]) {
      const read = bytes => (bigEndian ? Buffer.from(bytes.subarray(2)).swap16() : bytes.subarray(2)).toString("utf16le");
      assert.deepEqual([...output.get(part).subarray(0, 2)], bigEndian ? [254, 255] : [255, 254]);
      assert.equal(read(output.get(part)), read(parts.get(part)).replace(`>${before}</a:t>`, `>${after}</a:t>`));
    }
  } finally { session.dispose(); }
});
