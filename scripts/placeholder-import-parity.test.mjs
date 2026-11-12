import test from "node:test";
import assert from "node:assert/strict";
import { normalizePresentation, validatePresentation } from "../packages/core/dist/index.js";
import { resolveLayout } from "../packages/layout/dist/index.js";
import { importPptx, inspectOpcPackage } from "../packages/pptx-parser/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { renderPresentationToSvg } from "../packages/svg-renderer/dist/index.js";
import { placeholderBytes } from "../packages/pptx-parser/test/placeholder-fixture.mjs";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";

const semantics = document => normalizePresentation(document).slides.map(slide => slide.elements.map(({ id, ...element }) => element));

test("placeholder dimensions materialize exactly once over three fresh exports, with independently checked slide XML", async () => {
  let input = placeholderBytes(), expected;
  const sizes = [];
  for (let round = 0; round < 3; round++) {
    const session = await importPptx(input, { projection: "content" });
    assert.deepEqual(validatePresentation(session.document), []);
    expected ??= semantics(session.document);
    assert.deepEqual(semantics(session.document), expected);
    assert.deepEqual(resolveLayout(session.document).slides.map(slide => slide.elements.length), [1, 1]);
    if (round === 0) assert.equal(session.report.placeholderInheritance.placeholders.length, 4);
    else assert.equal(session.report.placeholderInheritance.placeholders.length, 0, "fresh output contains local values, not source template binding");
    const noOp = await session.prepareSave({ mode: "no-op" });
    assert.deepEqual((await session.save(noOp.plan)).bytes, new Uint8Array(input));
    const generated = await generatePptx(session.document), parts = readGeneratedZip(generated.bytes);
    assert.deepEqual(generated.warnings, []);
    for (let page = 1; page <= 2; page++) {
      const xml = parts.get(`ppt/slides/slide${page}.xml`).toString();
      assert.equal([...xml.matchAll(/<p:sp>/g)].length, 1);
      assert.match(xml, /<a:off x="127000" y="254000"\/>/);
      assert.match(xml, /<a:ext cx="2540000" cy="1270000"\/>/);
      assert.doesNotMatch(xml, /<p:ph|MASTER PROMPT/);
    }
    const opc = await inspectOpcPackage(generated.bytes);
    assert.equal(opc.relationships.filter(r => r.type.endsWith("/slideLayout") && r.source.startsWith("/ppt/slides/")).length, 2);
    const preview = await renderPresentationToSvg(session.document);
    assert.equal(preview.slides.length, 2);
    assert.match(preview.slides[0].svg, /A &amp; B/);
    assert.doesNotMatch(preview.slides[0].svg, /MASTER PROMPT/);
    sizes.push([...parts.values()].reduce((sum, value) => sum + value.length, 0));
    session.dispose(); input = generated.bytes;
  }
  assert.equal(new Set(sizes).size, 1, "no repeated inherited artwork or resource growth");
});

test("local edits after dimension projection leave shared-master pages unchanged and source save modes remain explicit", async () => {
  const session = await importPptx(placeholderBytes(), { projection: "content" }), original = semantics(session.document);
  const plan = await session.prepareSave({ mode: "no-op" });
  const slide = session.document.slides[0], previous = slide.removeElement(slide.elements[0].id);
  const changed = structuredClone(previous);
  changed.box = { ...changed.box, x: 0, width: 250 };
  changed.text.content = [{ runs: [{ text: "Local edit α😀", style: { color: "FF0000", bold: false } }] }];
  slide.addElement(changed);
  assert.deepEqual(semantics(session.document)[1], original[1]);
  await assert.rejects(session.save(plan.plan), e => e.code === "save-snapshot-stale");
  assert.equal((await session.prepareSave({ mode: "edit" })).status, "blocked");
  assert.equal((await session.prepareSave({ mode: "rebuild" })).status, "blocked");
  const fresh = await generatePptx(session.document), again = await importPptx(fresh.bytes, { projection: "content" });
  assert.equal(semantics(again.document)[0][0].box.x, 0);
  assert.equal(semantics(again.document)[0][0].box.width, 250);
  assert.equal(semantics(again.document)[0][0].text.plainText, "Local edit α😀");
  assert.deepEqual(semantics(again.document)[1], original[1]);
  again.dispose(); session.dispose();
});
