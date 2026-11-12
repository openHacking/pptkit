import assert from "node:assert/strict";
import test from "node:test";
import { createHash } from "node:crypto";
import { createPresentation } from "../packages/core/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { inspectOpcClosure, inspectOpcPackage, importPptx, compareOpcPackages } from "../packages/pptx-parser/dist/index.js";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";
import { zip } from "../packages/pptx-parser/test/zip-fixture.mjs";
import { MC } from "../packages/pptx-parser/test/closure-fixture.mjs";
import { png } from "../packages/pptx-parser/test/image-fixture.mjs";

const digest = (bytes) => createHash("sha256").update(bytes).digest("hex");
const repack = (parts) => zip([...parts].map(([name, body]) => ({ name, body }))).bytes;
async function sourceFixture() {
  const document = createPresentation();
  document.registerAsset({ id: "shared", kind: "image", source: { type: "bytes", value: png } });
  const a = document.addSlide(), b = document.addSlide();
  for (const slide of [a, b]) slide.addElement({ type: "image", assetId: "shared", box: { x: 10, y: 10, width: 80, height: 40 } });
  a.addElement({ type: "chart", chartType: "bar", categories: ["A", "B"], series: [{ name: "Series", values: [3, 4] }], box: { x: 100, y: 20, width: 300, height: 200 } });
  a.setNotes("Preservation dependency notes");
  const parts = readGeneratedZip((await generatePptx(document)).bytes);
  const relationshipPart = "ppt/slides/_rels/slide1.xml.rels";
  parts.set(relationshipPart, Buffer.from(parts.get(relationshipPart).toString()
    .replace("<Relationships ", `<Relationships xmlns:mc="${MC}" xmlns:x="urn:future" mc:Ignorable="x" `)
    .replace("</Relationships>", '<mc:AlternateContent><mc:Choice Requires="x"><Relationship Id="shadow" Type="urn:future-resource" Target="../media/shadow.bin"/></mc:Choice><mc:Fallback/></mc:AlternateContent></Relationships>')));
  parts.set("ppt/media/shadow.bin", Buffer.from([1, 5, 9]));
  parts.set("orphan.bin", Buffer.from([7, 8]));
  parts.set("[Content_Types].xml", Buffer.from(parts.get("[Content_Types].xml").toString()
    .replace("</Types>", '<Override PartName="/ppt/media/shadow.bin" ContentType="application/octet-stream"/><Override PartName="/orphan.bin" ContentType="application/octet-stream"/></Types>')));
  return repack(parts);
}

test("generated chart/notes/image source closure includes inactive dependencies, workbook bytes and original no-op state", async () => {
  const bytes = await sourceFixture(), parts = readGeneratedZip(bytes);
  const result = await inspectOpcClosure(bytes, { roots: ["/ppt/slides/slide1.xml"] });
  const selected = result.parts.map((p) => p.name);
  assert.ok(selected.includes("/ppt/media/shadow.bin"));
  assert.ok(selected.some((p) => p.startsWith("/ppt/embeddings/")));
  assert.ok(selected.includes("/ppt/notesSlides/notesSlide1.xml"));
  assert.ok(selected.includes("/ppt/slideMasters/slideMaster1.xml"));
  assert.ok(!selected.includes("/ppt/slides/slide2.xml"));
  assert.equal(result.sourceRelationshipCount, result.effectiveRelationshipCount + 1);
  assert.equal(result.relationships.find((r) => r.id === "shadow").matchesEffectiveRelationship, false);
  assert.deepEqual(result.coverage, { outgoingComplete: true, incomingComplete: true });
  const expectedBytes = result.parts.reduce((sum, part) => {
    assert.equal(part.expandedBytes, parts.get(part.entryName).length);
    return sum + parts.get(part.entryName).length;
  }, parts.get("[Content_Types].xml").length);
  assert.equal(result.expandedBytes, expectedBytes);
  const session = await importPptx(bytes, { projection: "content" });
  try {
    const ready = await session.prepareSave({ mode: "no-op" }); assert.equal(ready.status, "ready");
    const saved = (await session.save(ready.plan)).bytes;
    assert.equal(digest(saved), digest(bytes));
    assert.deepEqual(await inspectOpcClosure(saved, { roots: result.roots }), result);
  } finally { session.dispose(); }
});

test("shared image has an independently verifiable outside owner and unreachable data remains preserved", async () => {
  const bytes = await sourceFixture(), parts = readGeneratedZip(bytes);
  const result = await inspectOpcClosure(bytes, { roots: ["/ppt/slides/slide1.xml"] });
  const incoming = result.incomingRelationships.find((edge) => edge.source === "/ppt/slides/slide2.xml" && edge.type.endsWith("/image"));
  assert.ok(incoming);
  assert.equal(digest(parts.get(incoming.targetPart.slice(1))), digest(png));
  assert.ok(parts.get("ppt/slides/_rels/slide2.xml.rels").toString().includes(`Target="../media/${incoming.targetPart.split("/").at(-1)}"`));
  const packageClosure = await inspectOpcClosure(bytes, { roots: ["/"] });
  assert.ok(packageClosure.unreachedParts.includes("/orphan.bin"));
  assert.ok(packageClosure.parts.some((p) => p.name === "/ppt/media/shadow.bin"));
  const session = await importPptx(bytes, { projection: "content" });
  try {
    const ready = await session.prepareSave({ mode: "no-op" }); assert.equal(ready.status, "ready");
    const saved = readGeneratedZip((await session.save(ready.plan)).bytes);
    assert.equal(digest(saved.get("orphan.bin")), digest(parts.get("orphan.bin")));
    assert.equal((await session.prepareSave({ mode: "edit" })).status, "blocked");
    const fresh = await generatePptx(session.document);
    assert.equal((await compareOpcPackages(bytes, fresh.bytes)).equal, false);
  } finally { session.dispose(); }
});

test("hidden missing targets downgrade source coverage without being mistaken for active corruption or edit permission", async () => {
  const original = await sourceFixture(), parts = readGeneratedZip(original), name = "ppt/slides/_rels/slide1.xml.rels";
  parts.set(name, Buffer.from(parts.get(name).toString().replace('Target="../media/shadow.bin"', 'Target="../media/missing.bin"')));
  const brokenSource = repack(parts);
  assert.equal((await inspectOpcPackage(brokenSource)).relationships.length, (await inspectOpcPackage(original)).relationships.length);
  const closure = await inspectOpcClosure(brokenSource, { roots: ["/ppt/slides/slide1.xml"] });
  assert.equal(closure.coverage.outgoingComplete, false);
  assert.ok(closure.findings.some((f) => f.code === "missing-source-target" && f.source === "/ppt/slides/slide1.xml"));
  const session = await importPptx(brokenSource, { projection: "content" });
  try {
    const ready = await session.prepareSave({ mode: "no-op" }); assert.equal(ready.status, "ready");
    assert.equal(digest((await session.save(ready.plan)).bytes), digest(brokenSource));
    assert.equal((await session.prepareSave({ mode: "edit" })).status, "blocked");
  } finally { session.dispose(); }
});
