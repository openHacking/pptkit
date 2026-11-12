import assert from "node:assert/strict";
import test from "node:test";
import { createHash } from "node:crypto";
import { createPresentation } from "../packages/core/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { compareOpcPackages, importPptx } from "../packages/pptx-parser/dist/index.js";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";
import { zip } from "../packages/pptx-parser/test/zip-fixture.mjs";

const hash = (bytes) => createHash("sha256").update(bytes).digest("hex");
const repack = (parts, reverse = false) => zip([...parts].map(([name, body]) => ({ name, body, method: 0 })).sort((a, b) => reverse ? b.name.localeCompare(a.name) : a.name.localeCompare(b.name))).bytes;
async function createSource() {
  const p = createPresentation();
  const slide = p.addSlide();
  slide.addElement({ type: "text", box: { x: 10, y: 10, width: 300, height: 50 }, content: [{ runs: [{ text: "Comparison source", style: { color: "2468AC", colorOpacity: .25 } }] }] });
  slide.addElement({ type: "chart", chartType: "bar", box: { x: 10, y: 80, width: 300, height: 200 }, categories: ["A", "B"], series: [{ name: "Series", values: [2, 3] }] });
  slide.setNotes("Notes must survive");
  const parts = readGeneratedZip((await generatePptx(p)).bytes);
  parts.set("unknown/payload.bin", Buffer.from("plumless"));
  const types = parts.get("[Content_Types].xml").toString().replace("</Types>", '<Override PartName="/unknown/payload.bin" ContentType="application/octet-stream"/></Types>');
  parts.set("[Content_Types].xml", Buffer.from(types));
  const xml = parts.get("ppt/slides/slide1.xml").toString().replace("</p:sld>", '<p:extLst><p:ext uri="urn:test:opaque"><x:opaque xmlns:x="urn:test:unknown" keep="yes"/></p:ext></p:extLst></p:sld>');
  parts.set("ppt/slides/slide1.xml", Buffer.from(xml));
  return repack(parts);
}

test("independent ZIP/payload oracle validates original-byte sessions and ignores container differences", async () => {
  const source = await createSource(), parts = readGeneratedZip(source), reencoded = repack(parts, true);
  assert.notEqual(hash(source), hash(reencoded));
  const independent = readGeneratedZip(reencoded);
  assert.deepEqual([...parts.keys()].sort(), [...independent.keys()].sort());
  for (const [name, bytes] of parts) assert.equal(hash(independent.get(name)), hash(bytes), name);
  for (const profile of ["opc-payload-v1", "opc-xml-v1"]) assert.equal((await compareOpcPackages(source, reencoded, { profile })).equal, true);
  const session = await importPptx(source, { projection: "content" });
  try {
    const prepared = await session.prepareSave({ mode: "no-op" }); assert.equal(prepared.status, "ready");
    const saved = (await session.save(prepared.plan)).bytes;
    assert.equal(hash(saved), hash(source));
    assert.equal((await compareOpcPackages(source, saved)).equal, true);
    const fresh = await generatePptx(session.document), result = await compareOpcPackages(source, fresh.bytes);
    assert.equal(result.equal, false, "fresh output must not masquerade as preservation");
    assert.ok(result.differences.some((d) => d.code === "part-removed" && d.part === "/unknown/payload.bin"));
    assert.ok(result.differences.some((d) => d.part === "/ppt/slides/slide1.xml"));
  } finally { session.dispose(); }
});

test("independent single-fault candidate mutations cannot hide alpha, notes, unknown XML or workbook loss", async () => {
  const source = await createSource(), original = readGeneratedZip(source);
  const slide = "ppt/slides/slide1.xml", notes = "ppt/notesSlides/notesSlide1.xml";
  const workbook = [...original.keys()].find((name) => name.startsWith("ppt/embeddings/"));
  assert.ok(workbook);
  const mutations = [
    ["alpha", slide, (bytes) => Buffer.from(bytes.toString().replace('val="25000"', 'val="50000"'))],
    ["unknown extension", slide, (bytes) => Buffer.from(bytes.toString().replace('keep="yes"', 'keep="no"'))],
    ["notes", notes, (bytes) => Buffer.from(bytes.toString().replace("Notes must survive", "Changed notes"))],
    ["workbook", workbook, (bytes) => { const result = Buffer.from(bytes); result[result.length - 1] ^= 1; return result; }],
    ["equal-CRC binary", "unknown/payload.bin", () => Buffer.from("buckeroo")],
  ];
  for (const [label, part, mutate] of mutations) {
    const parts = new Map(original), changedBytes = mutate(parts.get(part));
    assert.notEqual(hash(changedBytes), hash(parts.get(part)), `${label} must actually mutate`);
    parts.set(part, changedBytes);
    const candidate = repack(parts);
    const independent = readGeneratedZip(candidate);
    assert.deepEqual([...original].filter(([name, bytes]) => hash(independent.get(name)) !== hash(bytes)).map(([name]) => name), [part]);
    for (const profile of ["opc-payload-v1", "opc-xml-v1"]) {
      const result = await compareOpcPackages(source, candidate, { profile });
      assert.equal(result.equal, false, label);
      assert.ok(result.differences.some((d) => d.part === `/${part}`), label);
    }
  }
  const missingWorkbook = new Map(original); missingWorkbook.delete(workbook);
  // Candidate validation catches dangling resources before any comparison result.
  await assert.rejects(compareOpcPackages(source, repack(missingWorkbook)), { code: "package-invalid" });
});

test("three independently reopened no-op rounds retain every source payload without growth", async () => {
  let bytes = await createSource();
  const original = readGeneratedZip(bytes), expectedSize = [...original.values()].reduce((sum, part) => sum + part.length, 0);
  for (let round = 0; round < 3; round++) {
    const session = await importPptx(bytes, { projection: "content" });
    try {
      const prepared = await session.prepareSave({ mode: "no-op" }); assert.equal(prepared.status, "ready");
      const saved = (await session.save(prepared.plan)).bytes;
      const result = await compareOpcPackages(bytes, saved, { profile: "opc-xml-v1" });
      assert.equal(result.equal, true);
      const parts = readGeneratedZip(saved);
      assert.equal(parts.size, original.size);
      assert.equal([...parts.values()].reduce((sum, part) => sum + part.length, 0), expectedSize);
      for (const [name, data] of original) assert.equal(hash(parts.get(name)), hash(data), name);
      bytes = saved;
    } finally { session.dispose(); }
  }
});
