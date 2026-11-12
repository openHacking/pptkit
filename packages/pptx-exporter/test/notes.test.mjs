import assert from "node:assert/strict";
import test from "node:test";
import { createPresentation } from "@pptkit/core";
import { generatePptx } from "../dist/index.js";
import { readGeneratedZip } from "../../../tests/support/read-generated-zip.mjs";

test("empty notes paragraphs retain authored style and actions rather than being dropped by text-length detection", async () => {
  const p = createPresentation(), s = p.addSlide({ id: "self" });
  s.setNotes([{ runs: [{ text: "", style: { fontSize: 12, bold: true }, action: { type: "slide", slideId: s.id, tooltip: "Self" } }], style: { align: "right" } }]);
  const result = await generatePptx(p), parts = readGeneratedZip(result.bytes);
  assert.deepEqual(result.warnings, []);
  const xml = parts.get("ppt/notesSlides/notesSlide1.xml").toString();
  assert.match(xml, /<p:ph type="body"/); assert.match(xml, /<a:pPr algn="r"/);
  assert.match(xml, /<a:rPr[^>]*sz="1200"[^>]*b="1"/); assert.match(xml, /ppaction:\/\/hlinksldjump/);
  assert.match(parts.get("ppt/notesSlides/_rels/notesSlide1.xml.rels").toString(), /Target="\.\.\/slides\/slide1\.xml"/);
});

test("an empty notes run with only explicit text alpha remains authored content", async () => {
  for (const alpha of [0, .5, 1]) {
    const p = createPresentation();
    p.addSlide({ notes: [{ runs: [{ text: "", style: { colorOpacity: alpha } }] }] });
    const { bytes } = await generatePptx(p), parts = readGeneratedZip(bytes);
    const xml = parts.get("ppt/notesSlides/notesSlide1.xml").toString();
    if (alpha === 1) assert.doesNotMatch(xml, /<p:ph type="body"/);
    else {
      assert.match(xml, /<p:ph type="body"/);
      assert.match(xml, new RegExp(`<a:alpha val="${alpha * 100000}"`));
    }
  }
});

test("empty paragraph insertion properties retain the final authored style without duplicating actions", async () => {
  const p = createPresentation(), s = p.addSlide({ id: "self" });
  s.setNotes([{ runs: [
    { text: "", style: { fontFamily: "Georgia", fontSize: 30 } },
    { text: "", style: { fontFamily: "Arial", eastAsianFontFamily: "Eastern", complexScriptFontFamily: "Complex", fontSize: 12, bold: true, italic: true, underline: true, strike: true, language: "fr-FR", color: "123456", colorOpacity: .5 }, action: { type: "slide", slideId: s.id } },
  ] }, { runs: [{ text: "Real", style: { fontSize: 22 } }] }]);
  const parts = readGeneratedZip((await generatePptx(p)).bytes), xml = parts.get("ppt/notesSlides/notesSlide1.xml").toString();
  const end = [...xml.matchAll(/<a:endParaRPr\b[^>]*>[\s\S]*?<\/a:endParaRPr>|<a:endParaRPr\b[^>]*\/>/g)].map(m => m[0]);
  assert.equal(end.length, 2);
  assert.match(end[0], /lang="fr-FR" sz="1200" b="1" i="1" u="sng" strike="sngStrike"/);
  for (const family of ["Arial", "Eastern", "Complex"]) assert.ok(end[0].includes(`typeface="${family}"`));
  assert.match(end[0], /val="123456"/); assert.match(end[0], /alpha val="50000"/);
  assert.doesNotMatch(end[0], /hlink/);
  assert.equal((xml.match(/ppaction:\/\/hlinksldjump/g) ?? []).length, 1);
  assert.equal(end[1], '<a:endParaRPr lang="en-US"/>');
});

test("the shared empty insertion serializer covers text boxes, shape text and table cells", async () => {
  const p = createPresentation(), s = p.addSlide(), content = [{ runs: [{ text: "", style: { fontFamily: "Georgia", fontSize: 13, italic: true, language: "el" } }] }];
  const box = { x: 10, y: 10, width: 120, height: 80 };
  s.addElement({ type: "text", box, content });
  s.addElement({ type: "shape", shape: "rect", box: { ...box, x: 140 }, text: { content } });
  s.addElement({ type: "table", box: { ...box, x: 270 }, columns: [120], rows: [{ cells: [{ content }] }] });
  const parts = readGeneratedZip((await generatePptx(p)).bytes), xml = parts.get("ppt/slides/slide1.xml").toString();
  const ends = [...xml.matchAll(/<a:endParaRPr\b[^>]*>[\s\S]*?<\/a:endParaRPr>/g)].map(m => m[0]);
  assert.equal(ends.length, 3);
  for (const end of ends) { assert.match(end, /lang="el" sz="1300" b="0" i="1" u="none" strike="noStrike"/); assert.match(end, /typeface="Georgia"/); }
});
