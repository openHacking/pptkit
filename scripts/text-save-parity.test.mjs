import assert from "node:assert/strict";
import test from "node:test";
import { createHash } from "node:crypto";
import { createPresentation, normalizePresentation } from "../packages/core/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { importPptx, compareOpcPackages, inspectOpcClosure } from "../packages/pptx-parser/dist/index.js";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";
import { zip } from "../packages/pptx-parser/test/zip-fixture.mjs";
import { png } from "../packages/pptx-parser/test/image-fixture.mjs";

const profile = { projection: "content", editSaveProfile: "text-runs-v1" };
const digest = bytes => createHash("sha256").update(bytes).digest("hex");
const repack = parts => zip([...parts].map(([name, body]) => ({ name, body }))).bytes;
const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
export async function textSaveSource() {
  const p = createPresentation(), first = p.addSlide(), second = p.addSlide();
  p.registerAsset({ id: "shared", kind: "image", source: { type: "bytes", value: png } });
  first.addElement({ type: "text", box: { x: 10, y: 10, width: 200, height: 60 }, content: [{ runs: [{ text: "Original text", style: { fontFamily: "Arial", color: "123456", colorOpacity: .65 }, action: { type: "slide", slideId: second.id } }] }] });
  first.addElement({ type: "group", box: { x: 300, y: 20, width: 180, height: 100 }, coordinateSize: { width: 90, height: 50 }, transform: { rotation: 15, flipH: true }, children: [
    { type: "text", box: { x: 5, y: 5, width: 80, height: 30 }, content: "Nested text" },
  ] });
  for (const slide of [first, second]) slide.addElement({ type: "image", assetId: "shared", box: { x: 10, y: 200, width: 80, height: 40 } });
  first.addElement({ type: "chart", chartType: "bar", categories: ["A", "B"], series: [{ name: "Data", values: [3, 4] }], box: { x: 100, y: 100, width: 200, height: 100 } });
  first.addElement({ type: "table", columns: [100, 100], rows: [{ height: 30, cells: [{ content: "A" }, { content: "B" }] }], box: { x: 100, y: 250, width: 200, height: 30 } });
  first.setNotes("Notes with untouched text");
  const parts = readGeneratedZip((await generatePptx(p)).bytes), slideName = "ppt/slides/slide1.xml";
  const opaque = '<p:ext uri="urn:future"><x:opaque xmlns:x="urn:future" ref="2">Original text</x:opaque></p:ext>';
  const slideXml = parts.get(slideName).toString();
  parts.set(slideName, Buffer.from(slideXml.includes("</p:extLst>") ? slideXml.replace("</p:extLst>", opaque + "</p:extLst>") : slideXml.replace("</p:sld>", `<p:extLst>${opaque}</p:extLst></p:sld>`)));
  const relName = "ppt/slides/_rels/slide1.xml.rels";
  parts.set(relName, Buffer.from(parts.get(relName).toString()
    .replace("<Relationships ", `<Relationships xmlns:mc="${MC}" xmlns:x="urn:future" mc:Ignorable="x" `)
    .replace("</Relationships>", '<mc:AlternateContent><mc:Choice Requires="x"><Relationship Id="inactive" Type="urn:future" Target="../media/shadow.bin"/></mc:Choice><mc:Fallback/></mc:AlternateContent></Relationships>')));
  parts.set("ppt/media/shadow.bin", Buffer.from([1, 3, 7])); parts.set("orphan.bin", Buffer.from([5, 9]));
  parts.set("[Content_Types].xml", Buffer.from(parts.get("[Content_Types].xml").toString().replace("</Types>", '<Override PartName="/ppt/media/shadow.bin" ContentType="application/octet-stream"/><Override PartName="/orphan.bin" ContentType="application/octet-stream"/></Types>')));
  return repack(parts);
}
function editFirst(session, text) {
  const slide = session.document.slides[0], index = slide.elements.findIndex(e => e.type === "text"), element = structuredClone(slide.removeElement(slide.elements[index].id));
  element.content[0].runs[0].text = text; slide.insertElement(index, element);
}
async function saved(session) {
  const preparation = await session.prepareSave({ mode: "edit" }); assert.equal(preparation.status, "ready", JSON.stringify(preparation.report));
  return session.save(preparation.plan);
}

test("text edit-save retains shared media, chart/workbook/table/notes and unknown source bytes", async () => {
  const input = await textSaveSource(), before = readGeneratedZip(input), closure = await inspectOpcClosure(input, { roots: ["/"] });
  const session = await importPptx(input, profile);
  try {
    editFirst(session, "Edited <&🙂");
    const groupIndex = session.document.slides[0].elements.findIndex(e => e.type === "group"), slide = session.document.slides[0];
    const group = structuredClone(slide.removeElement(slide.elements[groupIndex].id)); group.children[0].content[0].runs[0].text = "Edited nested"; slide.insertElement(groupIndex, group);
    const result = await saved(session), parts = readGeneratedZip(result.bytes), graph = await inspectOpcClosure(result.bytes, { roots: ["/"] });
    assert.equal(result.report.changes.length, 2);
    assert.ok(result.report.changes.some(c => c.documentPath.includes("/children/0/")));
    assert.deepEqual([...parts.keys()].sort(), [...before.keys()].sort());
    for (const [name, payload] of before) if (name !== "ppt/slides/slide1.xml") assert.equal(digest(parts.get(name)), digest(payload), name);
    assert.deepEqual(graph.relationships, closure.relationships); assert.deepEqual(graph.unreachedParts, closure.unreachedParts);
    const expected = before.get("ppt/slides/slide1.xml").toString().replace(">Original text</a:t>", ">Edited &lt;&amp;🙂</a:t>").replace(">Nested text</a:t>", ">Edited nested</a:t>");
    assert.equal(parts.get("ppt/slides/slide1.xml").toString(), expected);
    const comparison = await compareOpcPackages(input, result.bytes);
    assert.deepEqual(comparison.differences.map(d => [d.code, d.part]), [["payload-changed", "/ppt/slides/slide1.xml"]]);
    const reopened = await importPptx(result.bytes, profile);
    assert.deepEqual(normalizePresentation(reopened.document), normalizePresentation(session.document)); reopened.dispose();
  } finally { session.dispose(); }
});

test("production text save retains UTF-16BE slide bytes and inherited source encoding through three rounds", async () => {
  const parts = readGeneratedZip(await textSaveSource()), name = "ppt/slides/slide1.xml";
  const xml = parts.get(name).toString().replace(/encoding="UTF-8"/i, 'encoding="UTF-16"');
  parts.set(name, Buffer.concat([Buffer.from([254, 255]), Buffer.from(xml, "utf16le").swap16()]));
  let input = repack(parts), expectedXml = xml;
  for (let round = 1; round <= 3; round++) {
    const session = await importPptx(input, profile);
    try {
      editFirst(session, `UTF16 round ${round}🙂\r`);
      const result = await saved(session), output = readGeneratedZip(result.bytes).get(name);
      expectedXml = expectedXml.replace(round === 1 ? ">Original text</a:t>" : `>UTF16 round ${round - 1}🙂&#xD;</a:t>`, `>UTF16 round ${round}🙂&#xD;</a:t>`);
      assert.deepEqual(output, Buffer.concat([Buffer.from([254, 255]), Buffer.from(expectedXml, "utf16le").swap16()]));
      input = result.bytes;
    } finally { session.dispose(); }
  }
});

test("one unsupported edit blocks the entire mixed plan and restoring it permits exact text-only save", async () => {
  const input = await textSaveSource(), session = await importPptx(input, profile);
  try {
    editFirst(session, "Text proposal");
    const slide = session.document.slides[0], index = slide.elements.findIndex(e => e.type === "chart"), originalChart = slide.removeElement(slide.elements[index].id), modifiedChart = structuredClone(originalChart);
    modifiedChart.series[0].values[0] = 99; slide.insertElement(index, modifiedChart);
    const blocked = await session.prepareSave({ mode: "edit" }); assert.equal(blocked.status, "blocked"); assert.equal(blocked.report.assurance, "none");
    assert.ok(blocked.report.changes.some(c => c.kind === "text")); assert.ok(blocked.report.changes.some(c => c.kind !== "text"));
    slide.removeElement(modifiedChart.id); slide.insertElement(index, originalChart);
    const result = await saved(session), parts = readGeneratedZip(result.bytes), before = readGeneratedZip(input);
    for (const [name, payload] of parts) if (name !== "ppt/slides/slide1.xml") assert.equal(digest(payload), digest(before.get(name)), name);
    assert.equal(result.report.changes.length, 1);
  } finally { session.dispose(); }
});
