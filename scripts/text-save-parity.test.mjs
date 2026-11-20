import assert from "node:assert/strict";
import test from "node:test";
import { createHash } from "node:crypto";
import { readFile } from "node:fs/promises";
import { createPresentation, normalizePresentation } from "../packages/core/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { importPptx, compareOpcPackages, inspectOpcClosure, inspectPptx, inspectPptxTiming } from "../packages/pptx-parser/dist/index.js";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";
import { zip } from "../packages/pptx-parser/test/zip-fixture.mjs";
import { png } from "../packages/pptx-parser/test/image-fixture.mjs";
import { mixedAdvancedFixture } from "./research/mixed-advanced-fixture.mjs";

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

test("mixed text-save attribution detects resource, relationship and target corruption after three rounds", async () => {
  let bytes = await readFile("tests/fixtures/mixed-text-attribution-source.pptx"), previousText = "Original text";
  const evidence = JSON.parse(await readFile("docs/testing/mixed-text-attribution-evidence.json", "utf8"));
  assert.equal(digest(bytes), evidence.sourceSha256);
  assert.equal(evidence.rounds, 3); assert.deepEqual(evidence.detectedMutations, ["resource", "relationship", "target"]);
  const slidePart = "ppt/slides/slide1.xml", relPart = "ppt/slides/_rels/slide1.xml.rels", imagePart = "ppt/media/shared.png";
  for (let round = 1; round <= 3; round++) {
    const before = readGeneratedZip(bytes), session = await importPptx(bytes, profile), value = `Round ${round}`;
    try {
      editFirst(session, value);
      if (round === 1) {
        const slide = session.document.slides[0], index = slide.elements.findIndex(e => e.type === "chart"), original = slide.elements[index];
        const modified = structuredClone(original); modified.series[0].values[0] = 99; slide.replaceElement(original.id, modified);
        const blocked = await session.prepareSave({ mode: "edit" }); assert.equal(blocked.status, "blocked"); assert.equal(blocked.report.assurance, "none");
        slide.replaceElement(original.id, original);
      }
      const result = await saved(session), after = readGeneratedZip(result.bytes);
      assert.deepEqual([...after.keys()].sort(), [...before.keys()].sort());
      for (const [name, data] of before) {
        if (name === slidePart) {
          const original = data.toString(), anchor = `>${previousText}</a:t>`;
          assert.ok(original.includes(anchor));
          assert.equal(after.get(name).toString(), original.replace(anchor, `>${value}</a:t>`));
        } else assert.deepEqual(after.get(name), data, name);
      }
      const comparison = await compareOpcPackages(bytes, result.bytes);
      assert.deepEqual(comparison.differences.map(d => [d.code, d.part]), [["payload-changed", `/${slidePart}`]]);
      const reopened = await importPptx(result.bytes, profile);
      try { assert.equal(reopened.document.slides[0].elements.find(e => e.type === "text").content[0].runs[0].text, value); } finally { reopened.dispose(); }
      previousText = value; bytes = result.bytes;
    } finally { session.dispose(); }
  }
  const baseline = readGeneratedZip(bytes);
  for (const [label, part, mutate, code] of [
    ["resource", imagePart, body => { const changed = Buffer.from(body); changed[changed.length - 1] ^= 1; return changed; }, "payload-changed"],
    ["relationship", relPart, body => Buffer.from(body.toString().replace('Target="../media/shared.png"', 'Target="../media/shadow.bin"')), "relationship-changed"],
    ["target", slidePart, body => Buffer.from(body.toString().replace('<a:hlinkClick r:id="rId3"', '<a:hlinkClick r:id="rId2"')), "payload-changed"],
  ]) {
    const parts = new Map(baseline), original = parts.get(part), changed = mutate(original);
    assert.ok(original && !original.equals(changed), `${label} mutation must change bytes`);
    parts.set(part, changed);
    const candidate = repack(parts), independent = readGeneratedZip(candidate);
    assert.deepEqual([...baseline].filter(([name, data]) => !data.equals(independent.get(name))).map(([name]) => name), [part]);
    const comparison = await compareOpcPackages(bytes, candidate);
    assert.equal(comparison.equal, false, label);
    assert.ok(comparison.differences.some(d => d.code === code && d.part === `/${part}` && (label !== "relationship" || d.relationshipId === "rId4")), `${label}: ${JSON.stringify(comparison.differences)}`);
  }
});

test("mixed SmartArt and timing remain source-owned over three text saves", async () => {
  let bytes = await readFile("tests/fixtures/mixed-smartart-timing-source.pptx");
  const evidence = JSON.parse(await readFile("docs/testing/mixed-smartart-timing-evidence.json", "utf8"));
  const base = await readFile("tests/fixtures/mixed-text-attribution-source.pptx");
  assert.equal(digest(bytes), evidence.sourceSha256); assert.equal(digest(base), evidence.baseSha256);
  assert.equal(evidence.rounds, 3); assert.deepEqual(evidence.diagramComponents, ["data", "layout", "colors", "style", "cache"]);
  assert.equal(evidence.rangeEditBlocked, true); assert.equal(evidence.unrelatedRetrySaved, true); assert.equal(evidence.nativeApplication, "untested");
  assert.deepEqual(bytes, Buffer.from(mixedAdvancedFixture(base)));
  const initial = readGeneratedZip(bytes), baseline = await inspectPptx(bytes), timeline = await inspectPptxTiming(bytes);
  assert.equal(baseline.slides[0].diagrams.length, 1);
  assert.ok(baseline.slides[0].diagrams[0].components.every(c => c.status === "located"));
  assert.equal(baseline.slides[0].diagrams[0].cache.status, "located");
  assert.equal(baseline.diagramInventory.dataModels.length, 1); assert.equal(baseline.diagramInventory.cacheMappings.length, 1);
  assert.ok(timeline.slides[0].nodes.some(node => node.drawingReference?.rawId === "2" && node.drawingReference.status === "located"));
  const graph = await inspectOpcClosure(bytes, { roots: ["/"] }); let previous = "Original text";
  for (let round = 1; round <= 3; round++) {
    const before = readGeneratedZip(bytes), session = await importPptx(bytes, profile), next = `Advanced ${round}`;
    try {
      if (round === 1) { const plan = await session.prepareSave({ mode: "no-op" }); assert.deepEqual((await session.save(plan.plan)).bytes, new Uint8Array(bytes)); }
      editFirst(session, next);
      const output = (await saved(session)).bytes, after = readGeneratedZip(output), slide = "ppt/slides/slide1.xml";
      assert.deepEqual([...after.keys()].sort(), [...initial.keys()].sort());
      for (const [name, data] of before) {
        if (name === slide) {
          const xml = data.toString(), anchor = `>${previous}</a:t>`; assert.ok(xml.includes(anchor));
          assert.equal(after.get(name).toString(), xml.replace(anchor, `>${next}</a:t>`));
        } else assert.deepEqual(after.get(name), data, name);
      }
      assert.deepEqual((await inspectPptx(output)).diagramInventory, baseline.diagramInventory);
      assert.deepEqual(await inspectPptxTiming(output), timeline);
      const afterGraph = await inspectOpcClosure(output, { roots: ["/"] });
      assert.deepEqual(afterGraph.relationships, graph.relationships); assert.deepEqual(afterGraph.unreachedParts, graph.unreachedParts);
      assert.deepEqual((await compareOpcPackages(bytes, output)).differences.map(d => [d.code, d.part]), [["payload-changed", `/${slide}`]]);
      const reopened = await importPptx(output, profile);
      try { assert.equal(reopened.document.slides[0].elements.find(e => e.type === "text").content[0].runs[0].text, next); } finally { reopened.dispose(); }
      bytes = output; previous = next;
    } finally { session.dispose(); }
  }
  const unsafe = new Map(initial), part = "ppt/slides/slide1.xml", xml = unsafe.get(part).toString();
  assert.ok(xml.includes('<p:spTgt spid="2"/>'));
  unsafe.set(part, Buffer.from(xml.replace('<p:spTgt spid="2"/>', '<p:spTgt spid="2"><p:txEl><p:charRg st="0" end="2"/></p:txEl></p:spTgt>')));
  const session = await importPptx(repack(unsafe), profile);
  try {
    editFirst(session, "unsafe range edit");
    const plan = await session.prepareSave({ mode: "edit" }); assert.equal(plan.status, "blocked");
    assert.ok(plan.report.diagnostics.some(d => d.code === "save-text-range-reference"));
    editFirst(session, "Original text");
    const restored = await session.prepareSave({ mode: "no-op" }); assert.deepEqual((await session.save(restored.plan)).bytes, new Uint8Array(repack(unsafe)));
    const slide = session.document.slides[0], group = structuredClone(slide.elements.find(element => element.type === "group"));
    group.children[0].content[0].runs[0].text = "Unaffected nested text"; slide.replaceElement(group.id, group);
    const retried = readGeneratedZip((await saved(session)).bytes);
    assert.equal(retried.get(part).toString(), unsafe.get(part).toString().replace(">Nested text</a:t>", ">Unaffected nested text</a:t>"));
    for (const [name, data] of unsafe) if (name !== part) assert.deepEqual(retried.get(name), data, name);
  } finally { session.dispose(); }
});
