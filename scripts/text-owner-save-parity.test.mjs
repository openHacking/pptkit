import assert from "node:assert/strict";
import test from "node:test";
import { createHash } from "node:crypto";
import { readFile } from "node:fs/promises";
import { execFileSync } from "node:child_process";
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

test("native notes rewrite and subsequent three PPTKit edits have separate attribution", async () => {
  const evidence = JSON.parse(await readFile("docs/testing/corpus-notes-edit-powerpoint-evidence.json", "utf8"));
  const attribution = JSON.parse(await readFile("docs/testing/native-notes-attribution-evidence.json", "utf8"));
  const source = await readFile(evidence.source.path), native = await readFile(evidence.output.path);
  assert.equal(digest(source), evidence.source.sha256); assert.equal(digest(native), evidence.output.sha256);
  assert.equal(attribution.sourceSha256, digest(native)); assert.equal(attribution.nativeSourceSha256, digest(source));
  assert.match(execFileSync("python3", ["scripts/research/verify-corpus-notes-edit.py"], { encoding: "utf8" }), /verified notes edit package/);
  execFileSync("python3", ["scripts/research/test_corpus_notes_edit.py"], { encoding: "utf8", stdio: "pipe" });
  const expectedNative = ["docProps/app.xml", "docProps/core.xml", "ppt/notesSlides/notesSlide2.xml", "ppt/viewProps.xml"];
  assert.deepEqual(attribution.nativeChangedParts, expectedNative); assert.equal(attribution.pptkitRounds, 3);
  assert.equal(attribution.pptkitChangedPart, "ppt/notesSlides/notesSlide2.xml"); assert.equal(attribution.unsupportedFormattingBlocked, true);
  const beforeNative = readGeneratedZip(source), afterNative = readGeneratedZip(native);
  assert.deepEqual([...beforeNative.keys()].sort(), [...afterNative.keys()].sort());
  assert.deepEqual([...beforeNative].filter(([name, data]) => !data.equals(afterNative.get(name))).map(([name]) => name).sort(), expectedNative);
  const nativeComparison = await compareOpcPackages(source, native);
  assert.deepEqual(nativeComparison.differences.map(d => [d.code, d.part]), expectedNative.map(part => ["payload-changed", `/${part}`]));
  let input = native, previous = "native";
  for (let round = 1; round <= 3; round++) {
    const before = readGeneratedZip(input), session = await importPptx(input, profile), next = `PPTKit${round}`;
    try {
      const slide = session.document.slides[1], notes = structuredClone(slide.notes); assert.equal(notes[0].runs[2].text, previous);
      notes[0].runs[2].text = next; slide.setNotes(notes);
      if (round === 1) {
        const unsupported = structuredClone(notes); unsupported[0].runs[2].style.bold = true; slide.setNotes(unsupported);
        const blocked = await session.prepareSave({ mode: "edit" }); assert.equal(blocked.status, "blocked"); assert.equal(blocked.report.assurance, "none");
        slide.setNotes(notes);
      }
      const result = await save(session), after = readGeneratedZip(result.bytes), part = "ppt/notesSlides/notesSlide2.xml";
      assert.deepEqual([...after.keys()].sort(), [...before.keys()].sort());
      for (const [name, data] of before) {
        if (name === part) {
          const xml = data.toString(), anchor = `>${previous}</a:t>`; assert.ok(xml.includes(anchor));
          assert.equal(after.get(name).toString(), xml.replace(anchor, `>${next}</a:t>`));
        } else assert.deepEqual(after.get(name), data, name);
      }
      const comparison = await compareOpcPackages(input, result.bytes);
      assert.deepEqual(comparison.differences.map(d => [d.code, d.part]), [["payload-changed", `/${part}`]]);
      const reopened = await importPptx(result.bytes, profile);
      try { assert.equal(reopened.document.slides[1].notes[0].runs[2].text, next); } finally { reopened.dispose(); }
      input = result.bytes; previous = next;
    } finally { session.dispose(); }
  }
});

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
