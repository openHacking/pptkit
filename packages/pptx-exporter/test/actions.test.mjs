import assert from "node:assert/strict";
import test from "node:test";
import { createPresentation } from "@pptkit/core";
import { generatePptx } from "../dist/index.js";
import { readGeneratedZip } from "../../../tests/support/read-generated-zip.mjs";
import { pixelPng } from "../../../tests/fixtures/presentations.mjs";

test("internal slide actions carry explicit jump semantics at every currently modeled drawing/text location", async () => {
  const p = createPresentation(), action = { type: "slide", slideId: "target", tooltip: "Go <&>" }, text = [{ runs: [{ text: "Jump", action }] }];
  const s = p.addSlide({ id: "source", notes: text }); p.addSlide({ id: "target" });
  s.addElement({ type: "text", box: { x: 0, y: 0, width: 80, height: 30 }, content: text, action });
  s.addElement({ type: "shape", shape: "rect", box: { x: 0, y: 40, width: 80, height: 30 }, text: { content: text }, action });
  s.addElement({ type: "group", box: { x: 0, y: 80, width: 80, height: 30 }, coordinateSize: { width: 80, height: 30 }, children: [], action });
  s.addElement({ type: "connector", start: { x: 100, y: 10 }, end: { x: 180, y: 10 }, action });
  const asset = p.registerAsset({ kind: "image", source: { type: "url", value: pixelPng } });
  s.addElement({ type: "image", box: { x: 100, y: 40, width: 80, height: 30 }, assetId: asset.id, action });
  s.addElement({ type: "table", box: { x: 100, y: 80, width: 80, height: 30 }, columns: [80], rows: [{ height: 30, cells: [{ content: text }] }], action });
  const result = await generatePptx(p, { connectorMode: "native" }), parts = readGeneratedZip(result.bytes);
  assert.deepEqual(result.warnings, []);
  for (const [part, count, target] of [["ppt/slides/slide1.xml", 9, "slide2.xml"], ["ppt/notesSlides/notesSlide1.xml", 1, "../slides/slide2.xml"]]) {
    const tags = [...parts.get(part).toString().matchAll(/<a:hlinkClick\b[^>]*\/>/g)].map(m => m[0]);
    assert.equal(tags.length, count);
    const relPart = part.replace(/\/([^/]+)$/, "/_rels/$1.rels"), rels = parts.get(relPart).toString();
    for (const tag of tags) {
      assert.match(tag, /action="ppaction:\/\/hlinksldjump"/);
      assert.match(tag, /tooltip="Go &lt;&amp;&gt;"/);
      const id = /r:id="([^"]+)"/.exec(tag)[1];
      const relationship = [...rels.matchAll(/<Relationship\b[^>]+\/>/g)].map(m => m[0]).find(tag => tag.includes(`Id="${id}"`));
      assert.ok(relationship); assert.match(relationship, /Type="[^"]+\/slide"/);
      assert.equal(/Target="([^"]+)"/.exec(relationship)[1], target); assert.doesNotMatch(relationship, /TargetMode="External"/);
    }
  }
});

test("external URL actions keep hyperlink relationships and never acquire internal slide-jump commands", async () => {
  const p = createPresentation(), s = p.addSlide(), action = { type: "url", url: "https://example.invalid/a?x=1&y=%E4%B8%AD" };
  s.addElement({ type: "text", box: { x: 0, y: 0, width: 100, height: 40 }, content: [{ runs: [{ text: "URL", action }] }], action });
  const parts = readGeneratedZip((await generatePptx(p)).bytes), xml = parts.get("ppt/slides/slide1.xml").toString(), rels = parts.get("ppt/slides/_rels/slide1.xml.rels").toString();
  assert.equal((xml.match(/<a:hlinkClick\b/g) ?? []).length, 2);
  assert.doesNotMatch(xml, /ppaction:/);
  assert.equal((rels.match(/TargetMode="External"/g) ?? []).length, 2);
  assert.match(rels, /Target="https:\/\/example.invalid\/a\?x=1&amp;y=%E4%B8%AD"/);
});
