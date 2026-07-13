import test from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, readFile, stat, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { inflateRawSync } from "node:zlib";

import { createPresentation } from "@pptkit/core";
import { generatePptx } from "../dist/index.js";
import { generatePptx as generateNodePptx, writePptx } from "../dist/node.js";

function readZipEntries(input) {
  const bytes = Buffer.from(input);
  const entries = new Map();
  let offset = 0;
  while (offset + 30 <= bytes.length && bytes.readUInt32LE(offset) === 0x04034b50) {
    const nameLength = bytes.readUInt16LE(offset + 26);
    const extraLength = bytes.readUInt16LE(offset + 28);
    const compressedSize = bytes.readUInt32LE(offset + 18);
    const method = bytes.readUInt16LE(offset + 8);
    const name = bytes.subarray(offset + 30, offset + 30 + nameLength).toString();
    const data = bytes.subarray(offset + 30 + nameLength + extraLength, offset + 30 + nameLength + extraLength + compressedSize);
    entries.set(name, method === 8 ? inflateRawSync(data) : data);
    offset += 30 + nameLength + extraLength + compressedSize;
  }
  return entries;
}

test("generatePptx exports the IR v1 feature surface", async () => {
  const presentation = createPresentation({
    metadata: { title: "Q1 & <ready>", author: "PPTKit", company: "Example", language: "zh-CN" },
    theme: { colors: { accent1: "123456" }, fonts: { heading: "Arial", body: "Microsoft YaHei" } },
  });
  presentation.defineSlideLayout({
    id: "title-content",
    name: "Title & Content",
    background: { type: "solid", color: { theme: "background2" } },
    elements: [{ type: "shape", shape: "rect", box: { x: 0, y: 0, width: 12, height: 540 }, style: { fill: { type: "solid", color: { theme: "accent1" } } } }],
    placeholders: [{ key: "title", kind: "title", box: { x: 40, y: 30, width: 700, height: 60 }, textStyle: { run: { fontSize: 36, fontFamily: { theme: "heading" } } } }],
  });
  const asset = presentation.registerAsset({ kind: "image", source: { type: "url", value: "https://example.com/hero.png" }, width: 400, height: 200, accessibility: { description: "Hero" } });
  const first = presentation.addSlide({
    id: "intro",
    layoutId: "title-content",
    notes: "Speaker note",
    section: "Opening",
    tags: ["hero"],
    customData: { source: "unit-test" },
  });
  first.addElement({
    type: "text",
    content: [{
      style: { bullet: { type: "bullet", character: "→" }, lineSpacing: 1.15 },
      runs: [
        { text: "Q1 & <ready> ", style: { bold: true, color: { theme: "accent1" } }, action: { type: "url", url: "https://example.com", tooltip: "Visit" } },
        { text: "next", style: { italic: true }, action: { type: "slide", slideId: "details" } },
      ],
    }],
    placeholderKey: "title",
    frame: { verticalAlign: "middle", autoFit: { mode: "shrink", fontScale: 0.95 } },
  });
  first.addElement({ type: "shape", id: "left", shape: "roundRect", box: { x: 40, y: 130, width: 120, height: 60 }, style: { fill: { type: "solid", color: "FF0000", opacity: 0.8 } } });
  first.addElement({ type: "shape", id: "right", shape: "ellipse", box: { x: 300, y: 130, width: 80, height: 80 } });
  first.addElement({ type: "connector", start: { elementId: "left", anchor: "right" }, end: { elementId: "right", anchor: "left" }, route: [{ x: 220, y: 150 }], style: { dash: "dash", endArrow: "triangle" } });
  first.addElement({
    type: "group",
    box: { x: 420, y: 120, width: 240, height: 140 },
    coordinateSize: { width: 240, height: 140 },
    children: [{ type: "image", assetId: asset.id, fit: "cover", box: { x: 0, y: 0, width: 240, height: 140 } }],
  });
  first.addElement({
    type: "table",
    box: { x: 40, y: 300, width: 400, height: 120 },
    columns: [200, 200],
    rows: [
      { height: 50, cells: [{ content: "Header", colSpan: 2, style: { fill: { type: "solid", color: { theme: "accent1" } } } }] },
      { height: 70, cells: [{ content: "A" }, { content: "B" }] },
    ],
  });
  presentation.addSlide({ id: "details", hidden: true, background: { type: "solid", color: "FAFAFA" }, elements: [{ type: "text", content: "Details", box: { x: 40, y: 40, width: 300, height: 50 } }] });

  const originalFetch = globalThis.fetch;
  globalThis.fetch = async () => new Response(Buffer.from("fake-png"), { status: 200, headers: { "content-type": "image/png" } });
  try {
    const result = await generatePptx(presentation);
    const repeated = await generatePptx(presentation);
    assert.deepEqual(result.bytes, repeated.bytes);
    assert.equal(result.status, "generated");
    assert.equal(result.slideCount, 2);
    const entries = readZipEntries(result.bytes);
    const slide1 = entries.get("ppt/slides/slide1.xml").toString();
    const slide2 = entries.get("ppt/slides/slide2.xml").toString();
    const slideRels = entries.get("ppt/slides/_rels/slide1.xml.rels").toString();
    const layout2 = entries.get("ppt/slideLayouts/slideLayout2.xml").toString();
    const theme = entries.get("ppt/theme/theme1.xml").toString();
    const notes = entries.get("ppt/notesSlides/notesSlide1.xml").toString();

    assert.match(entries.get("docProps/core.xml").toString(), /Q1 &amp; &lt;ready&gt;/);
    assert.match(theme, /accent1><a:srgbClr val="123456"/);
    assert.match(theme, /majorFont><a:latin typeface="Arial"/);
    assert.match(theme, /minorFont><a:latin typeface="Microsoft YaHei"/);
    assert.match(layout2, /name="Title &amp; Content"/);
    assert.match(layout2, /<p:ph type="title" idx="1"/);
    assert.match(layout2, /<a:schemeClr val="accent1"/);
    assert.doesNotMatch(slide1, /<p:bg>/);
    assert.match(slide1, /Q1 &amp; &lt;ready&gt;/);
    assert.match(slide1, /<a:buChar char="→"/);
    assert.match(slide1, /<a:normAutofit fontScale="95000"/);
    assert.match(slide1, /<p:cxnSp>/);
    assert.match(slide1, /<a:custGeom>/);
    assert.match(slide1, /<p:grpSp>/);
    assert.match(slide1, /<a:srcRect l="7143"[^>]*r="7143"/);
    assert.match(slide1, /<a:tbl>/);
    assert.match(slide1, /gridSpan="2"/);
    assert.match(slide1, /hMerge="1"/);
    assert.match(slide1, /pptkit:slideData/);
    assert.match(slideRels, /TargetMode="External"/);
    assert.match(slideRels, /relationships\/hyperlink/);
    assert.match(slideRels, /Target="\.\.\/slides\/slide2.xml"/);
    assert.match(notes, /Speaker note/);
    assert.match(slide2, /show="0"/);
    assert.match(slide2, /<p:bg>/);
    assert.ok(entries.has("ppt/media/asset-1.png"));
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test("Node entry packages local images and writes nested output", async () => {
  const directory = await mkdtemp(join(tmpdir(), "pptkit-export-"));
  const imagePath = join(directory, "pixel.png");
  await writeFile(imagePath, Buffer.from("fake-png"));
  const presentation = createPresentation();
  const asset = presentation.registerAsset({ kind: "image", id: "hero", source: { type: "path", value: imagePath }, mimeType: "image/png" });
  const missing = presentation.registerAsset({ kind: "image", id: "missing", source: { type: "path", value: join(directory, "missing.png") }, mimeType: "image/png" });
  presentation.addSlide({ elements: [
    { type: "image", assetId: asset.id, box: { x: 0, y: 0, width: 100, height: 100 } },
    { type: "image", assetId: missing.id, box: { x: 100, y: 0, width: 100, height: 100 } },
  ] });
  const output = join(directory, "deck.pptx");
  const generated = await generateNodePptx(presentation);
  const result = await writePptx(presentation, { output });
  assert.equal(result.status, "written-with-warnings");
  assert.ok(result.warnings.some((warning) => warning.code === "asset-read-failed"));
  assert.ok(readZipEntries(generated.bytes).has("ppt/media/hero.png"));
  assert.equal((await stat(output)).size, result.byteLength);
  assert.ok(readZipEntries(await readFile(output)).has("ppt/media/hero.png"));
});

test("default entry reports path assets as unsupported", async () => {
  const presentation = createPresentation();
  const asset = presentation.registerAsset({ kind: "image", source: { type: "path", value: "./pixel.png" } });
  presentation.addSlide({ elements: [{ type: "image", assetId: asset.id, box: { x: 0, y: 0, width: 100, height: 100 } }] });
  const result = await generatePptx(presentation);
  assert.equal(result.status, "generated-with-warnings");
  assert.match(result.warnings.find((warning) => warning.code === "asset-read-failed")?.message ?? "", /only supported by @pptkit\/pptx-exporter\/node/);
});
