import test from "node:test";
import assert from "node:assert/strict";
import { orientJpegCrop, generatePptx } from "../dist/index.js";
import { readFile } from "node:fs/promises";
import { unzipSync, strFromU8 } from "fflate";
import { createPresentation } from "@pptkit/core";
import { resolveLayout } from "@pptkit/layout";
import { orientJpegCrop as nodeCrop } from "../dist/node.js";
import { orientJpegCrop as domCrop } from "../dist/dom.js";

test("EXIF crop mapping preserves asymmetric regions for all orientations", () => {
  const source = Object.freeze({ left: 0.1, top: 0.2, right: 0.3, bottom: 0.4 });
  const expected = [[.1,.2,.3,.4],[.3,.2,.1,.4],[.3,.4,.1,.2],[.1,.4,.3,.2],
    [.2,.1,.4,.3],[.4,.1,.2,.3],[.4,.3,.2,.1],[.2,.3,.4,.1]];
  const inverses = [1,2,3,4,5,8,7,6];
  for (let orientation = 1; orientation <= 8; orientation++) {
    const mapped = orientJpegCrop(source, orientation);
    assert.deepEqual([mapped.left,mapped.top,mapped.right,mapped.bottom], expected[orientation-1]);
    assert.notEqual(mapped, source);
    assert.deepEqual(orientJpegCrop(mapped, inverses[orientation-1]), source);
    assert.deepEqual(nodeCrop(source, orientation), mapped);
    assert.deepEqual(domCrop(source, orientation), mapped);
  }
});

test("EXIF crop mapping rejects invalid regions and orientation", () => {
  const crop = { left: 0, top: 0, right: 0, bottom: 0 };
  assert.deepEqual(orientJpegCrop(crop, 6), crop);
  for (const invalid of [-1, 1.01, Infinity, NaN, undefined, "0"]) {
    assert.throws(() => orientJpegCrop({ ...crop, left: invalid }, 1), RangeError);
  }
  assert.throws(() => orientJpegCrop({ ...crop, left: .5, right: .5 }, 1), RangeError);
  assert.throws(() => orientJpegCrop({ ...crop, top: .6, bottom: .5 }, 1), RangeError);
  for (const orientation of [0, 9, 1.5, NaN, "6", undefined]) assert.throws(() => orientJpegCrop(crop, orientation), RangeError);
  assert.throws(() => orientJpegCrop(null, 1), TypeError);
});


test("mapped raw crop and oriented dimensions survive layout and PPTX serialization", async () => {
  const png = await readFile(new URL("../../../tests/fixtures/exif-orientation6-normalized.png", import.meta.url));
  const deck = createPresentation();
  const asset = deck.registerAsset({ id: "oriented", kind: "image", source: { type: "bytes", value: png }, mimeType: "image/png", width: 2, height: 3 });
  const crop = orientJpegCrop({ left: 0, top: 0, right: 2/3, bottom: 1/2 }, 6);
  deck.addSlide({ elements: [{ type: "image", id: "cropped", assetId: asset.id, fit: "crop", crop, box: { x: 0, y: 0, width: 120, height: 120 } }] });
  assert.deepEqual(resolveLayout(deck).slides[0].elements[0].crop, crop);
  const result = await generatePptx(deck);
  const zip = unzipSync(result.bytes);
  const xml = strFromU8(zip["ppt/slides/slide1.xml"]);
  assert.match(xml, /<a:srcRect l="50000" t="0" r="0" b="66667"\/>/);
  const media = Object.entries(zip).filter(([name]) => name.startsWith("ppt/media/"));
  assert.equal(media.length, 1);
  assert.deepEqual(media[0][1], new Uint8Array(png));
});
