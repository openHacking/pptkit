import test from "node:test";
import assert from "node:assert/strict";
import { normalizePresentation, validatePresentation } from "../packages/core/dist/index.js";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { textStyleFixture, setTextTheme, setLayoutThemeFont, layoutList } from "../packages/pptx-parser/test/text-style-fixture.mjs";
import { encodeStructure } from "../packages/pptx-parser/test/structure-fixture.mjs";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";
import { createPresentation } from "../packages/core/dist/index.js";
import { renderPresentationToSvg } from "../packages/svg-renderer/dist/index.js";

const semantics = document => normalizePresentation(document).slides.map(slide => slide.elements.map(({ id, ...element }) => element));
const sourceBytes = () => {
  const f = setLayoutThemeFont(setTextTheme(textStyleFixture(), { accent1: "2468AC", major: "Base Parity Font" }), "Parity Theme Font");
  f.parts["deck/layout.xml"] = f.parts["deck/layout.xml"].replace(layoutList,
    '<a:lvl1pPr><a:defRPr sz="2000"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:latin typeface="+mj-lt"/></a:defRPr></a:lvl1pPr>');
  return encodeStructure(f);
};

test("text inheritance materializes stable local values over three fresh export/import rounds", async () => {
  let input = sourceBytes(), expected;
  const sizes = [];
  for (let round = 0; round < 3; round++) {
    const session = await importPptx(input, { projection: "content" });
    assert.deepEqual(validatePresentation(session.document), []);
    expected ??= semantics(session.document);
    assert.deepEqual(semantics(session.document), expected);
    const [shape] = session.document.slides[0].elements;
    const text = shape.text;
    assert.deepEqual(text.frame.margin, { left: 5, right: 2, top: 0, bottom: 4 });
    assert.equal(text.content[0].runs[0].style.fontSize, 20);
    assert.equal(text.content[0].runs[0].style.color, "2468AC");
    assert.equal(text.content[0].runs[0].style.fontFamily, "Parity Theme Font");
    assert.equal(text.content[0].style.align, "center");
    assert.deepEqual(text.content[0].style.bullet, { type: "bullet", character: "*" });
    assert.doesNotMatch(JSON.stringify(semantics(session.document)), /MASTER PROMPT/);

    if (round === 0) {
      assert.ok(session.report.textStyleInheritance.properties.length > 0);
      assert.ok(session.report.textStyleInheritance.properties.some(p => p.level === "layout-placeholder" && p.property === "run.fontSize"));
      assert.ok(session.report.textStyleInheritance.properties.some(p => p.property === "run.color" && p.resolvedFrom?.some(source => source.part === "/deck/theme.xml")));
      assert.ok(session.report.textStyleInheritance.properties.some(p => p.property === "run.fontFamily" && p.resolvedFrom?.some(source => source.part === "/deck/font-override.xml")));
    } else {
      assert.ok(session.report.textStyleInheritance.properties.every(p => p.level === "run" || p.level === "paragraph" || p.level === "end-paragraph" || p.level === "slide-frame" || p.level === "slide-list-style" || (p.level === "presentation-default" && p.property === "paragraph.direction" && p.value === "ltr")));
    }

    const noOp = await session.prepareSave({ mode: "no-op" });
    assert.equal(noOp.status, "ready");
    assert.ok(Buffer.from((await session.save(noOp.plan)).bytes).equals(Buffer.from(input)));
    const generated = await generatePptx(session.document), parts = readGeneratedZip(generated.bytes);
    assert.deepEqual(generated.warnings, []);
    for (let page = 1; page <= 2; page++) {
      const xml = parts.get(`ppt/slides/slide${page}.xml`).toString();
      const body = /<a:bodyPr\b[^>]*>/.exec(xml)?.[0] ?? "";
      assert.match(body, /lIns="63500"/); assert.match(body, /rIns="25400"/);
      assert.match(body, /tIns="0"/); assert.match(body, /bIns="50800"/);
      assert.match(xml, /<a:rPr\b[^>]*sz="2000"[^>]*b="1"[^>]*u="sng"/);
      assert.match(xml, /<a:srgbClr val="2468AC"/);
      assert.match(xml, /<a:latin typeface="Parity Theme Font"/);
      assert.match(xml, /<a:pPr[^>]*algn="ctr"/);
      assert.doesNotMatch(xml, /<p:ph|MASTER PROMPT MUST NOT COPY/);
    }
    sizes.push([...parts.values()].reduce((sum, value) => sum + value.length, 0));
    session.dispose();
    input = generated.bytes;
  }
  assert.equal(new Set(sizes).size, 1, "materialized styles do not grow on repeated fresh exports");
});

test("editing one page leaves shared-template styles on the other page unchanged", async () => {
  const session = await importPptx(sourceBytes(), { projection: "content" });
  const before = semantics(session.document);
  const original = session.document.slides[0].removeElement(session.document.slides[0].elements[0].id);
  const edited = structuredClone(original);
  edited.text.content[0].runs[0].text = "Edited locally";
  edited.text.content[0].runs[0].style.fontSize = 42;
  session.document.slides[0].addElement(edited);
  assert.deepEqual(semantics(session.document)[1], before[1]);

  const generated = await generatePptx(session.document);
  const reopened = await importPptx(generated.bytes, { projection: "content" });
  const after = semantics(reopened.document);
  assert.equal(after[0][0].text.content[0].runs[0].style.fontSize, 42);
  assert.deepEqual(after[1], before[1]);
  assert.equal((await session.prepareSave({ mode: "edit" })).status, "blocked");
  assert.equal((await session.prepareSave({ mode: "rebuild" })).status, "blocked");
  reopened.dispose();
  session.dispose();
});

test("selected text-style provenance enforces its independent session limit", async () => {
  await assert.rejects(importPptx(sourceBytes(), {
    projection: "content",
    sessionLimits: { maxTextStyleProperties: 1 },
  }), error => error.code === "import-limit-exceeded");
});

test("text alpha reaches native SVG, mixed-run HTML, PPTX and three fresh import/export rounds", async () => {
  const p = createPresentation({ size: { width: 400, height: 200 } }), slide = p.addSlide();
  const runs = [.25, 0, 1].map((colorOpacity, i) => ({ text: `Alpha ${i}`, style: { color: "336699", colorOpacity, fontFamily: "Liberation Sans" } }));
  slide.addElement({ type: "text", id: "native", box: { x: 0, y: 0, width: 200, height: 50 }, content: [{ runs: [runs[0]] }] });
  slide.addElement({ type: "text", id: "mixed", box: { x: 0, y: 50, width: 300, height: 50 }, content: [{ runs }] });
  slide.addElement({ type: "shape", id: "shape", shape: "rect", box: { x: 0, y: 100, width: 150, height: 50 }, text: { content: [{ runs: [runs[0]] }] } });
  slide.addElement({ type: "table", id: "table", box: { x: 200, y: 100, width: 150, height: 50 }, columns: [150], rows: [{ cells: [{ content: [{ runs: [runs[0]] }] }] }] });
  slide.setNotes([{ runs: [runs[0]] }]);
  const preview = await renderPresentationToSvg(p);
  assert.match(preview.slides[0].svg, /fill-opacity="0.25"/);
  assert.match(preview.slides[0].svg, /opacity:0.25/);
  assert.match(preview.slides[0].svg, /opacity:0(?:;|")/);
  const before = normalizePresentation(p);
  let input = (await generatePptx(p)).bytes;
  for (let round = 0; round < 3; round++) {
    const parts = readGeneratedZip(input), xml = parts.get("ppt/slides/slide1.xml").toString();
    assert.match(xml, /<a:alpha val="25000"/);
    assert.match(xml, /<a:alpha val="0"/);
    assert.match(parts.get("ppt/notesSlides/notesSlide1.xml").toString(), /<a:alpha val="25000"/);
    const s = await importPptx(input, { projection: "content" });
    try {
      assert.deepEqual(validatePresentation(s.document), []);
      const [native, mixed, shape, table] = s.document.slides[0].elements;
      assert.equal(native.content[0].runs[0].style.colorOpacity, .25);
      assert.deepEqual(mixed.content[0].runs.map(r => r.style.colorOpacity), [.25, 0, 1]);
      assert.equal(shape.text.content[0].runs[0].style.colorOpacity, .25);
      assert.equal(table.rows[0].cells[0].content[0].runs[0].style.colorOpacity, .25);
      assert.equal(s.document.slides[0].notes[0].runs[0].style.colorOpacity, .25);
      const noOp = await s.prepareSave({ mode: "no-op" });
      assert.deepEqual((await s.save(noOp.plan)).bytes, new Uint8Array(input));
      input = (await generatePptx(s.document)).bytes;
    } finally { s.dispose(); }
  }
  assert.deepEqual(normalizePresentation(p), before);
});

test("element alpha and text paint alpha multiply once without changing fills or opaque siblings", async () => {
  const p = createPresentation();
  p.addSlide().addElement({ type: "shape", shape: "rect", opacity: .5, style: { fill: { type: "solid", color: "FF0000" } },
    box: { x: 0, y: 0, width: 200, height: 80 }, text: { content: [{ runs: [
      { text: "Half of half", style: { color: "0000FF", colorOpacity: .5 } },
      { text: "Element only", style: { color: "0000FF" } },
    ] }] } });
  const { bytes } = await generatePptx(p), xml = readGeneratedZip(bytes).get("ppt/slides/slide1.xml").toString();
  assert.match(xml, /<a:srgbClr val="FF0000"><a:alpha val="50000"/);
  assert.match(xml, /<a:srgbClr val="0000FF"><a:alpha val="25000"/);
  assert.match(xml, /<a:srgbClr val="0000FF"><a:alpha val="50000"/);
  const preview = await renderPresentationToSvg(p);
  assert.match(preview.slides[0].svg, /opacity="0.5"/);
  assert.match(preview.slides[0].svg, /opacity:0.5/);
});

test("inherited scheme transforms materialize once and remain stable through three fresh rounds", async () => {
  const f = setTextTheme(textStyleFixture(), { accent1: "123456" });
  f.parts["deck/layout.xml"] = f.parts["deck/layout.xml"].replace(layoutList,
    '<a:lvl1pPr><a:defRPr><a:solidFill><a:schemeClr val="accent1"><a:lumMod val="50000"/><a:alpha val="25000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl1pPr>');
  let input = encodeStructure(f), expected;
  const sizes = [];
  for (let round = 0; round < 3; round++) {
    const s = await importPptx(input, { projection: "content" });
    try {
      const run = s.document.slides[0].elements[0].text.content[0].runs[0];
      assert.equal(run.style.color, "091A2B");
      assert.equal(run.style.colorOpacity, .25);
      expected ??= semantics(s.document);
      assert.deepEqual(semantics(s.document), expected);
      const generated = await generatePptx(s.document), parts = readGeneratedZip(generated.bytes);
      assert.match(parts.get("ppt/slides/slide1.xml").toString(), /<a:srgbClr val="091A2B"><a:alpha val="25000"\/>/);
      assert.doesNotMatch(parts.get("ppt/slides/slide1.xml").toString(), /lumMod/);
      sizes.push([...parts.values()].reduce((sum, bytes) => sum + bytes.length, 0));
      input = generated.bytes;
    } finally { s.dispose(); }
  }
  assert.equal(new Set(sizes).size, 1);
});
