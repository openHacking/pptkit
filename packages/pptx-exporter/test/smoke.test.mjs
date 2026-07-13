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

test("generatePptx creates deterministic cross-runtime bytes", async () => {
  const presentation = createPresentation({ title: "Quarterly plan" });
  presentation.addSlide({ elements: [
    { type: "text", text: "Q1 & <ready>", box: { x: 10, y: 20, width: 300, height: 40 }, style: { fontSize: 20, fontWeight: "bold" } },
    { type: "shape", shape: "rect", box: { x: 10, y: 70, width: 100, height: 50 }, style: { fill: "#ff0000" } },
  ] });
  const result = await generatePptx(presentation);
  const repeated = await generatePptx(presentation);

  assert.ok(result.bytes instanceof Uint8Array);
  assert.equal(result.slideCount, 1);
  assert.equal(result.status, "generated");
  assert.equal(result.warnings.length, 0);
  assert.equal(result.byteLength, result.bytes.byteLength);
  assert.deepEqual(result.bytes, repeated.bytes);
  const entries = readZipEntries(result.bytes);
  assert.match(entries.get("ppt/slides/slide1.xml").toString(), /Q1 &amp; &lt;ready&gt;/);
  const presentationXml = entries.get("ppt/presentation.xml").toString();
  const slideXml = entries.get("ppt/slides/slide1.xml").toString();
  const masterXml = entries.get("ppt/slideMasters/slideMaster1.xml").toString();
  assert.match(presentationXml, /sldId/);
  assert.match(presentationXml, /notesMasterIdLst/);
  assert.ok(presentationXml.indexOf("notesMasterIdLst") < presentationXml.indexOf("sldIdLst"));
  assert.match(slideXml, /algn="l"/);
  assert.doesNotMatch(slideXml, /algn="left"/);
  assert.match(masterXml, /accent3="accent3"/);
  assert.match(masterXml, /id="2147483649"/);
  assert.match(masterXml, /<p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"\/><\/p:bgRef><\/p:bg>/);
  assert.match(masterXml, /<p:titleStyle>/);
  assert.match(masterXml, /<p:bodyStyle>/);
  assert.match(masterXml, /<a:grpSpPr>|<p:grpSpPr>/);
  assert.match(slideXml, /<p:grpSpPr><a:xfrm>/);
  assert.match(entries.get("ppt/_rels/presentation.xml.rels").toString(), /relationships\/notesMaster/);
  assert.ok(entries.has("ppt/notesMasters/notesMaster1.xml"));
  assert.ok(entries.has("ppt/notesSlides/notesSlide1.xml"));
  assert.ok(entries.has("ppt/presProps.xml"));
  assert.ok(entries.has("ppt/viewProps.xml"));
  assert.ok(entries.has("ppt/tableStyles.xml"));
  assert.ok(entries.has("ppt/theme/theme2.xml"));
  assert.match(entries.get("ppt/notesMasters/_rels/notesMaster1.xml.rels").toString(), /Target="\.\.\/theme\/theme2\.xml"/);
  const themeXml = entries.get("ppt/theme/theme1.xml").toString();
  assert.match(themeXml, /<a:majorFont><a:latin[^>]*\/><a:ea[^>]*\/><a:cs[^>]*\/><\/a:majorFont>/);
  assert.match(themeXml, /<a:fillStyleLst>(?:.|\n)*<a:solidFill>(?:.|\n)*<\/a:fillStyleLst>/);
  assert.equal((themeXml.match(/<a:solidFill>/g) ?? []).length, 9);
  assert.ok(entries.has("[Content_Types].xml"));
});

test("Node entry packages local images and writes nested output", async () => {
  const directory = await mkdtemp(join(tmpdir(), "pptkit-export-"));
  const imagePath = join(directory, "pixel.png");
  await writeFile(imagePath, Buffer.from("fake-png"));
  const presentation = createPresentation();
  const asset = presentation.registerAsset({ kind: "image", id: "hero", source: { type: "path", value: imagePath }, mimeType: "image/png" });
  const missingAsset = presentation.registerAsset({ kind: "image", id: "missing", source: { type: "path", value: join(directory, "missing.png") }, mimeType: "image/png" });
  presentation.addSlide({ elements: [
    { type: "image", assetId: asset.id, box: { x: 0, y: 0, width: 100, height: 100 } },
    { type: "image", assetId: missingAsset.id, box: { x: 100, y: 0, width: 100, height: 100 } },
  ] });
  const output = join(directory, "deck.pptx");

  const generated = await generateNodePptx(presentation);
  const result = await writePptx(presentation, { output });

  assert.equal(result.status, "written-with-warnings");
  assert.ok(result.warnings.some((warning) => warning.code === "asset-read-failed"));
  assert.ok(readZipEntries(generated.bytes).has("ppt/media/hero.png"));
  const entries = readZipEntries(await readFile(output));
  assert.ok(entries.has("ppt/media/hero.png"));
  assert.match(entries.get("ppt/slides/_rels/slide1.xml.rels").toString(), /relationships\/image/);
  assert.equal((await stat(output)).size, result.byteLength);
});

test("generatePptx loads URL image assets", async () => {
  const presentation = createPresentation();
  const asset = presentation.registerAsset({ kind: "image", id: "remote", source: { type: "url", value: "http://127.0.0.1:1/asset.png" }, mimeType: "image/png" });
  presentation.addSlide({ elements: [{ type: "image", assetId: asset.id, box: { x: 0, y: 0, width: 100, height: 100 } }] });
  const originalFetch = globalThis.fetch;
  globalThis.fetch = async () => new Response(Buffer.from("remote-png"), { status: 200, headers: { "content-type": "image/png" } });
  try {
    const result = await generatePptx(presentation);
    assert.equal(result.status, "generated");
    assert.ok(readZipEntries(result.bytes).has("ppt/media/remote.png"));
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test("default entry reports path assets as unsupported", async () => {
  const presentation = createPresentation();
  const asset = presentation.registerAsset({ kind: "image", id: "local", source: { type: "path", value: "./pixel.png" } });
  presentation.addSlide({ elements: [{ type: "image", assetId: asset.id, box: { x: 0, y: 0, width: 100, height: 100 } }] });

  const result = await generatePptx(presentation);

  assert.equal(result.status, "generated-with-warnings");
  assert.match(result.warnings.find((warning) => warning.code === "asset-read-failed")?.message ?? "", /only supported by @pptkit\/pptx-exporter\/node/);
});
