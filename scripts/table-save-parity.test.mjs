import assert from "node:assert/strict";
import test from "node:test";
import { createHash } from "node:crypto";
import { createPresentation, normalizePresentation } from "../packages/core/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { importPptx, compareOpcPackages, inspectOpcClosure } from "../packages/pptx-parser/dist/index.js";
import { createTableSaveFixture } from "../tests/fixtures/table-save.mjs";
import { png } from "../packages/pptx-parser/test/image-fixture.mjs";
import { zip } from "../packages/pptx-parser/test/zip-fixture.mjs";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";

const profile = { projection: "content", editSaveProfile: "text-runs-v3" };
const hash = bytes => createHash("sha256").update(bytes).digest("hex");
async function source() {
  const parts = readGeneratedZip((await generatePptx(createTableSaveFixture(createPresentation, png))).bytes);
  const name = "ppt/slides/slide1.xml", xml = parts.get(name).toString();
  parts.set(name, Buffer.from(xml.replace("<a:tcPr", '<a:tcPr xmlns:x="urn:future" x:opaque="cell-control"')));
  parts.set("orphan.xml", Buffer.from('<future xmlns="urn:future">retained</future>'));
  return zip([...parts].map(([name, body]) => ({ name, body }))).bytes;
}
function edit(session, round) {
  const slide = session.document.slides[0];
  for (const element of slide.elements) {
    const copy = structuredClone(element);
    if (copy.type === "shape" && copy.text) copy.text.content[0].runs[0].text = `Shape round ${round}`;
    else if (copy.type === "table" && copy.columns.length === 2) copy.rows[0].cells[0].content[0].runs[0].text = `Cell round ${round}`;
    else if (copy.type === "table" && copy.columns.length === 3) {
      copy.rows[1].cells[0].content[0].runs[0].text = `Merged round ${round}`;
      copy.rows[2].cells[0].content[0].runs[0].text = `Last row round ${round}`;
    } else continue;
    const index = slide.elements.indexOf(element); slide.removeElement(element.id); slide.insertElement(index, copy);
  }
  const notes = structuredClone(slide.notes); notes[0].runs[0].text = `Notes round ${round}`; slide.setNotes(notes);
}
async function save(session) {
  const plan = await session.prepareSave({ mode: "edit" }); assert.equal(plan.status, "ready", JSON.stringify(plan.report));
  return session.save(plan.plan);
}

test("three mixed-owner v3 rounds retain chart/workbook/images/unknown parts and both table grids", async () => {
  let input = await source(), previousSize;
  const original = readGeneratedZip(input), graph = await inspectOpcClosure(input, { roots: ["/"] });
  for (let round = 1; round <= 3; round++) {
    const session = await importPptx(input, profile); edit(session, round);
    const saved = await save(session), parts = readGeneratedZip(saved.bytes);
    assert.equal(saved.report.changes.length, 5); assert.deepEqual(saved.report.losses, []);
    const changes = (await compareOpcPackages(input, saved.bytes)).differences;
    assert.deepEqual(changes.map(d => [d.code, d.part]), [["payload-changed", "/ppt/notesSlides/notesSlide1.xml"], ["payload-changed", "/ppt/slides/slide1.xml"]]);
    for (const [name, bytes] of original) if (!["ppt/slides/slide1.xml", "ppt/notesSlides/notesSlide1.xml"].includes(name)) assert.equal(hash(parts.get(name)), hash(bytes), name);
    assert.ok(parts.get("ppt/slides/slide1.xml").toString().includes('x:opaque="cell-control"'));
    const currentGraph = await inspectOpcClosure(saved.bytes, { roots: ["/"] });
    assert.deepEqual(currentGraph.relationships, graph.relationships);
    const reopened = await importPptx(saved.bytes, profile);
    assert.deepEqual(normalizePresentation(reopened.document), normalizePresentation(session.document)); reopened.dispose();
    const merged = session.report.tableTextRuns.filter(run => run.columnIndex === 2 && run.rowIndex === 1);
    assert.equal(merged.length, 1); assert.equal(merged[0].cellIndex, 0);
    if (previousSize) assert.equal(saved.bytes.length, previousSize);
    previousSize = saved.bytes.length; input = saved.bytes; session.dispose();
  }
});

test("v3 table/style mixtures block atomically and prior owner profiles remain compatible", async () => {
  const input = await source(), session = await importPptx(input, profile);
  edit(session, 1);
  const slide = session.document.slides[0], original = slide.elements.find(e => e.type === "table"), index = slide.elements.indexOf(original);
  const copy = structuredClone(original); copy.columns[0] += 10;
  slide.removeElement(original.id); slide.insertElement(index, copy);
  assert.equal((await session.prepareSave({ mode: "edit" })).status, "blocked");
  slide.removeElement(original.id); slide.insertElement(index, original); await save(session); session.dispose();
  for (const editSaveProfile of ["text-runs-v1", "text-runs-v2"]) {
    const prior = await importPptx(input, { projection: "content", editSaveProfile });
    assert.equal(prior.report.tableTextRuns, undefined); edit(prior, 2);
    assert.equal((await prior.prepareSave({ mode: "edit" })).status, "blocked"); prior.dispose();
  }
});
