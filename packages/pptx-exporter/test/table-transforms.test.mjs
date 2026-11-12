import assert from "node:assert/strict";
import test from "node:test";
import { createPresentation } from "@pptkit/core";
import { generatePptx } from "../dist/index.js";
import { readGeneratedZip } from "../../../tests/support/read-generated-zip.mjs";

test("native editable table groups carry rotation/flips while graphic frames stay neutral and shape namespaces stay intact", async () => {
  for (const transform of [{ rotation: 45, flipH: true, flipV: true }, { rotation: -30, flipV: true }, {}]) {
    const p = createPresentation(), slide = p.addSlide();
    slide.addElement({ type: "table", box: { x: 10, y: 20, width: 100, height: 50 }, transform, columns: [40, 60], rows: [{ height: 50, cells: [{ content: "Left" }, { content: "Right" }] }] });
    slide.addElement({ type: "shape", shape: "rect", box: { x: 30, y: 40, width: 20, height: 30 }, transform: { rotation: 15 } });
    const result = await generatePptx(p), xml = readGeneratedZip(result.bytes).get("ppt/slides/slide1.xml").toString();
    const grouped = Boolean(transform.rotation || transform.flipH || transform.flipV);
    assert.deepEqual(result.warnings.map(w => w.code), grouped ? ["table-transform-grouped"] : []);
    const attrs = `${transform.rotation ? ` rot="${transform.rotation * 60000}"` : ""}${transform.flipH ? ' flipH="1"' : ""}${transform.flipV ? ' flipV="1"' : ""}`;
    if (grouped) {
      assert.ok(xml.includes(`<a:xfrm${attrs}><a:off x="127000" y="254000"/><a:ext cx="1270000" cy="635000"/><a:chOff x="0" y="0"/><a:chExt cx="1270000" cy="635000"/></a:xfrm>`));
      assert.ok(xml.includes('<p:xfrm><a:off x="0" y="0"/><a:ext cx="1270000" cy="635000"/></p:xfrm>'));
      assert.ok(xml.includes('<p:grpSp>'));
    } else assert.ok(xml.includes('<p:xfrm><a:off x="127000" y="254000"/><a:ext cx="1270000" cy="635000"/></p:xfrm>'));
    assert.match(xml, /<p:spPr><a:xfrm rot="900000">/);
    assert.equal((xml.match(/<p:xfrm/g) ?? []).length, 1);
  }
});

test("grouped table identity, action and opacity remain singular and do not collide with following drawings", async () => {
  const document = createPresentation(), slide = document.addSlide();
  slide.addElement({ id: "table", type: "table", name: "Authored table", accessibility: { description: "Table description" }, action: { type: "url", url: "https://example.com/table" }, opacity: .5, transform: { rotation: 15 }, box: { x: 10, y: 20, width: 100, height: 50 }, columns: [100], rows: [{ height: 50, cells: [{ content: "Control", style: { stroke: { paint: { type: "solid", color: "FF0000" } }, borders: {} } }] }] });
  slide.addElement({ id: "shape", type: "shape", shape: "rect", box: { x: 0, y: 0, width: 5, height: 5 } });
  const xml = readGeneratedZip((await generatePptx(document)).bytes).get("ppt/slides/slide1.xml").toString();
  const ids = [...xml.matchAll(/<p:cNvPr id="(\d+)"/g)].map(m => m[1]); assert.equal(new Set(ids).size, ids.length);
  assert.match(xml, /<p:nvGrpSpPr><p:cNvPr id="2" name="Authored table" descr="Table description"><a:hlinkClick/);
  assert.equal((xml.match(/<a:hlinkClick/g) ?? []).length, 1);
  const lines = [...xml.matchAll(/<a:ln[LRBT]\b[^>]*>([\s\S]*?)<\/a:ln[LRBT]>/g)];
  assert.equal(lines.length, 4);
  for (const line of lines) assert.match(line[1], /<a:alpha val="50000"/);
  assert.equal(xml.includes('val="25000"'), false);
});
