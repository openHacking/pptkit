import assert from "node:assert/strict";
import test from "node:test";
import { readFile, mkdtemp, writeFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { normalizePresentation, validatePresentation } from "@pptkit/core";
import { importPptx } from "../dist/index.js";
import { importPptxFile } from "../dist/node.js";
import { picture, imageFixture, encodeImages, png, gif, jpegHeader } from "./image-fixture.mjs";
import { shape, transform } from "./content-fixture.mjs";
import { relationships } from "./opc-fixture.mjs";

const open = (f = imageFixture(), options = {}) => importPptx(encodeImages(f), { projection: "content", ...options });
const first = s => normalizePresentation(s.document).slides[0].elements[0];
const findings = s => s.report.findings.filter(f => f.operation === "import");

test("native H.264 video remains a diagnosed static poster and preserves the source", async () => {
  const bytes = await readFile(new URL("../../../tests/fixtures/h264-mp4-powerpoint.pptx", import.meta.url));
  const session = await importPptx(bytes, { projection: "content" });
  try {
    const diagnostic = session.report.diagnostics.find(x => x.code === "media-playback-unmapped");
    assert.equal(diagnostic.source.sourceObjectId, "4");
    assert.match(diagnostic.message, /static image is a media poster only/);
    assert.equal(session.document.slides[0].elements.find(x => x.id === diagnostic.elementId).type, "image");
    const plan = await session.prepareSave({ mode: "no-op" });
    assert.deepEqual((await session.save(plan.plan)).bytes, new Uint8Array(bytes));
    assert.equal((await session.prepareSave({ mode: "edit" })).status, "blocked");
  } finally { session.dispose(); }
});

test("opt-in media placeholders preserve source geometry and bytes without enabling edits", async () => {
  for (const strict of [false, true]) for (const fields of [{ embed: null }, { preset: "ellipse" }]) {
    const media = picture({ id: 2, ...fields }).replace('<p:nvPr/>', '<p:nvPr><a:audioFile r:link="media"/></p:nvPr>');
    const f = imageFixture(media + picture({ id: 3 }), strict), bytes = encodeImages(f);
    const s = await importPptx(bytes, { projection: "content", mediaPreview: "placeholder-v1" });
    try {
      const [placeholder, ordinary] = s.document.slides[0].elements;
      assert.equal(placeholder.type, "shape"); assert.equal(ordinary.type, "image");
      assert.deepEqual(placeholder.box, { x: 10, y: 20, width: 100, height: 50 });
      assert.equal(placeholder.text.content, "Media");
      assert.equal(s.report.findings.find(x => x.feature === "drawing.media.placeholder").tier, "fallback");
      assert.equal(s.report.findings.find(x => x.feature === "drawing.media.poster").tier, "fallback");
      assert.equal(s.report.findings.find(x => x.feature === "drawing.media" && x.operation === "preview").tier, "unsupported");
      assert.equal(s.report.diagnostics.find(x => x.code === "media-playback-unmapped").elementId, placeholder.id);
      assert.equal((await s.prepareSave({ mode: "edit" })).status, "blocked");
      const plan = await s.prepareSave({ mode: "no-op" }); assert.deepEqual((await s.save(plan.plan)).bytes, new Uint8Array(bytes));
      assert.deepEqual(validatePresentation(s.document), []);
      const modified = structuredClone(placeholder); modified.box.x += 1;
      s.document.slides[0].removeElement(placeholder.id); s.document.slides[0].insertElement(0, modified);
      assert.equal((await s.prepareSave({ mode: "no-op" })).status, "blocked");
    } finally { s.dispose(); }
  }
  const absent = picture({ embed: null, xfrm: "" }).replace('<p:nvPr/>', '<p:nvPr><a:audioFile r:link="media"/></p:nvPr>');
  const s = await open(imageFixture(absent), { mediaPreview: "placeholder-v1" });
  assert.equal(s.document.slides[0].elements.length, 0);
  assert.equal(s.report.findings.find(x => x.feature === "drawing.media.poster").tier, "unsupported"); s.dispose();
  await assert.rejects(open(imageFixture(), { mediaPreview: "placeholder-v1", editSaveProfile: "text-runs-v1" }), TypeError);
  await assert.rejects(open(imageFixture(), { mediaPreview: "invalid" }), TypeError);
  await assert.rejects(open(imageFixture(), { projection: "structure", mediaPreview: "placeholder-v1" }), TypeError);
});

test("file import captures the media placeholder option before asynchronous I/O", async () => {
  const directory = await mkdtemp(join(tmpdir(), "pptkit-media-preview-")), file = join(directory, "source.pptx");
  try {
    const xml = picture({ embed: null }).replace('<p:nvPr/>', '<p:nvPr><a:videoFile r:link="external"/></p:nvPr>');
    await writeFile(file, encodeImages(imageFixture(xml)));
    const options = { projection: "content", mediaPreview: "placeholder-v1" }, pending = importPptxFile(file, options);
    options.mediaPreview = undefined;
    const { session } = await pending;
    try { assert.equal(session.document.slides[0].elements[0].type, "shape"); } finally { session.dispose(); }
  } finally { await rm(directory, { recursive: true, force: true }); }
});

test("media pictures expose static poster fallback without claiming playback in both dialects", async () => {
  for (const strict of [false, true]) for (const kind of ["audioFile", "videoFile", "quickTimeFile", "audioCd", "wavAudioFile"]) {
    const media = picture({ id: 2 }).replace('<p:nvPr/>', `<p:nvPr><a:${kind} r:link="media"/></p:nvPr>`);
    const f = imageFixture(media + picture({ id: 3 }), strict);
    f.parts["deck/_rels/z.xml.rels"] = relationships([f.rel("layout", "slideLayout", "layout.xml"), f.rel("image", "image", "../media/raster.bin"), { ...f.rel("media", "audio", "https://example.invalid/media"), TargetMode: "External" }]);
    const bytes = encodeImages(f), s = await importPptx(bytes, { projection: "content" });
    try {
      const [poster, ordinary] = s.document.slides[0].elements;
      assert.equal(poster.type, "image"); assert.equal(ordinary.type, "image");
      const diagnostic = s.report.diagnostics.filter(item => item.code === "media-playback-unmapped");
      assert.equal(diagnostic.length, 1); assert.equal(diagnostic[0].elementId, poster.id); assert.match(diagnostic[0].source.path, new RegExp(`${kind}\\[1\\]$`));
      const mediaRows = s.report.findings.filter(item => item.elementId === poster.id && item.feature === "drawing.media");
      assert.equal(mediaRows.length, 7); assert.equal(mediaRows.find(item => item.operation === "preview").tier, "unsupported");
      assert.equal(mediaRows.find(item => item.operation === "no-op-save").tier, "preserved");
      assert.equal(s.report.findings.find(item => item.elementId === poster.id && item.feature === "drawing.media.poster").tier, "fallback");
      assert.equal(s.report.findings.find(item => item.elementId === poster.id && item.feature === "drawing.image.mapped-fields" && item.operation === "preview").tier, "fallback");
      assert.equal(s.report.findings.find(item => item.elementId === ordinary.id && item.feature === "drawing.image.mapped-fields" && item.operation === "preview").tier, "editable");
      const plan = await s.prepareSave({ mode: "no-op" }); assert.deepEqual((await s.save(plan.plan)).bytes, new Uint8Array(bytes));
    } finally { s.dispose(); }
  }
});

test("media without a mapped poster remains source-only and reports its actual source", async () => {
  for (const strict of [false, true]) for (const fields of [{ embed: null }, { xfrm: "" }, { preset: "ellipse" }]) {
    const media = picture({ id: 2, ...fields }).replace('<p:nvPr/>', '<p:nvPr><a:audioFile r:link="media"/></p:nvPr>');
    const f = imageFixture(media + picture({ id: 3 }), strict), bytes = encodeImages(f), s = await importPptx(bytes, { projection: "content" });
    try {
      assert.equal(s.document.slides[0].elements.length, 1); assert.equal(s.document.slides[0].elements[0].name, "Picture 3");
      const diagnostic = s.report.diagnostics.filter(item => item.code === "media-playback-unmapped");
      assert.equal(diagnostic.length, 1); assert.equal(diagnostic[0].elementId, undefined);
      assert.equal(diagnostic[0].source.sourceObjectId, "2"); assert.match(diagnostic[0].source.path, /audioFile\[1\]$/);
      assert.match(diagnostic[0].message, /no mapped static poster/);
      const poster = s.report.findings.find(item => item.feature === "drawing.media.poster");
      assert.equal(poster.tier, "unsupported"); assert.equal(poster.elementId, undefined);
      assert.equal(s.report.findings.find(item => item.feature === "drawing.media" && item.operation === "no-op-save").tier, "preserved");
      const plan = await s.prepareSave({ mode: "no-op" }); assert.deepEqual((await s.save(plan.plan)).bytes, new Uint8Array(bytes));
    } finally { s.dispose(); }
  }
});

test("Office media extension is detected only at its specified nonvisual path and URI", async () => {
  const guid = "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}";
  for (const strict of [false, true]) for (const [uri, namespace, mapped, expected] of [
    [guid, "http://schemas.microsoft.com/office/powerpoint/2010/main", true, true],
    [guid.toLowerCase(), "http://schemas.microsoft.com/office/powerpoint/2010/main", false, true],
    ["{00000000-0000-0000-0000-000000000000}", "http://schemas.microsoft.com/office/powerpoint/2010/main", true, false],
    [guid, "urn:other", true, false],
  ]) {
    const ext = `<p:extLst><p:ext uri="${uri}"><m:media xmlns:m="${namespace}" r:embed="media"/></p:ext></p:extLst>`;
    const xml = picture({ embed: mapped ? "image" : null }).replace('<p:nvPr/>', `<p:nvPr>${ext}</p:nvPr>`);
    const s = await open(imageFixture(xml, strict));
    try {
      const diagnostics = s.report.diagnostics.filter(item => item.code === "media-playback-unmapped");
      assert.equal(diagnostics.length, expected ? 1 : 0);
      if (expected) {
        assert.match(diagnostics[0].source.path, /media\[1\]$/);
        const preview = s.report.findings.find(item => item.feature === "drawing.media.poster");
        assert.equal(preview.tier, mapped ? "fallback" : "unsupported");
      }
    } finally { s.dispose(); }
  }
});

test("inactive MCE media declarations do not classify an ordinary effective picture", async () => {
  const xml = picture().replace('<p:nvPr/>', '<p:nvPr><mc:AlternateContent><mc:Choice Requires="x"><a:audioFile r:link="unused"/></mc:Choice><mc:Fallback/></mc:AlternateContent></p:nvPr>');
  const f = imageFixture(xml);
  f.parts["deck/z.xml"] = f.parts["deck/z.xml"].replace('<p:sld ', '<p:sld xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future-media" mc:Ignorable="x" ');
  const s = await open(f);
  try {
    assert.equal(s.report.diagnostics.filter(item => item.code === "media-playback-unmapped").length, 0);
    assert.equal(s.report.findings.some(item => item.feature === "drawing.media"), false);
    assert.equal(s.report.findings.find(item => item.feature === "drawing.image.mapped-fields" && item.operation === "preview").tier, "editable");
  } finally { s.dispose(); }
});

test("embedded images import exact immutable bytes, cropped geometry, transforms and opacity in both dialects", async () => {
  for (const strict of [false, true]) {
    const f = imageFixture(picture({ crop: '<a:srcRect l="10000" r="20000" t="12500"/>', xfrm: transform('rot="2700000" flipH="true" flipV="1"'), nv: 'hidden="1"', effects: '<a:alphaModFix amt="65000"/>' }), strict);
    const s = await open(f), e = first(s), asset = s.document.assets[0];
    assert.equal(e.type, "image"); assert.equal(e.fit, "crop"); assert.equal(e.opacity, .65); assert.equal(e.hidden, true);
    assert.deepEqual(e.box, { x: 10, y: 20, width: 100, height: 50 });
    assert.deepEqual(e.transform, { rotation: 45, flipH: true, flipV: true });
    assert.deepEqual(e.crop, { left: .1, right: .2, top: .125, bottom: 0 });
    assert.deepEqual(e.accessibility, { description: "Raster", decorative: false });
    assert.equal(e.assetId, asset.id); assert.equal(asset.mimeType, "image/png");
    assert.deepEqual(asset.source, { type: "bytes", value: png.toString("base64"), encoding: "base64" });
    assert.equal(asset.width, 8); assert.equal(asset.height, 4);
    assert.deepEqual(validatePresentation(s.document), []);
    const ids = s.report.identities.filter(i => i.kind === "asset");
    assert.equal(ids.length, 2); assert.equal(ids[0].source.part, "/media/raster.bin");
    assert.equal(ids[1].source.relationshipId, "image"); assert.equal(ids[1].source.sourceObjectId, "2");
    assert.ok(Object.isFrozen(ids[1].source));
    assert.equal(findings(s).filter(f => f.source.path?.endsWith("/@r:embed")).length, 0);
    assert.equal(findings(s).find(f => f.feature === "drawing.image.mapped-fields").tier, "editable");
    assert.ok(findings(s).some(f => /header-classified only/.test(f.reason))); s.dispose();
  }
});

test("shared targets resolve by canonical source part across relationships, groups and slides", async () => {
  const group = `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="4" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm flipH="1"><a:off x="0" y="0"/><a:ext cx="2540000" cy="1270000"/><a:chOff x="63500" y="88900"/><a:chExt cx="1270000" cy="635000"/></a:xfrm></p:grpSpPr>${picture({ id: 5, embed: "alias" })}</p:grpSp>`;
  const f = imageFixture(picture() + group);
  f.parts["deck/_rels/z.xml.rels"] = relationships([f.rel("layout", "slideLayout", "layout.xml"), f.rel("image", "image", "../media/raster.bin"), f.rel("alias", "image", "../MEDIA/RASTER.BIN")]);
  f.parts["deck/a.xml"] = f.parts["deck/z.xml"];
  f.parts["deck/_rels/a.xml.rels"] = f.parts["deck/_rels/z.xml.rels"];
  const s = await open(f, { sessionLimits: { maxImportedAssetBytes: png.length, maxTotalImportedAssetBytes: png.length } });
  assert.equal(s.document.assets.length, 1);
  const slides = normalizePresentation(s.document).slides;
  assert.deepEqual(slides[0].elements[1].children[0].box, { x: 5, y: 13, width: 100, height: 50 });
  const mapped = slides.flatMap(s => [s.elements[0], s.elements[1].children[0]]);
  assert.equal(new Set(mapped.map(e => e.id)).size, 4); assert.equal(new Set(mapped.map(e => e.assetId)).size, 1);
  assert.equal(s.report.identities.filter(i => i.kind === "asset").length, 5); s.dispose();
});

test("equal bytes in distinct source parts retain independent asset identity and budget charges", async () => {
  const f = imageFixture(picture() + picture({ id: 3, embed: "second" }));
  f.parts["media/second.png"] = png; f.mediaTypes["media/second.png"] = "image/png";
  f.parts["deck/_rels/z.xml.rels"] = relationships([f.rel("layout", "slideLayout", "layout.xml"), f.rel("image", "image", "../media/raster.bin"), f.rel("second", "image", "../media/second.png")]);
  await assert.rejects(open(f, { sessionLimits: { maxTotalImportedAssetBytes: png.length * 2 - 1 } }), { code: "import-limit-exceeded" });
  const s = await open(f, { sessionLimits: { maxTotalImportedAssetBytes: png.length * 2 } });
  assert.equal(s.document.assets.length, 2); assert.notEqual(s.document.slides[0].elements[0].assetId, s.document.slides[0].elements[1].assetId); s.dispose();
});

test("external pictures never become live Core URL assets and embedded-plus-linked pictures report update loss", async () => {
  for (const embedded of [false, true]) {
    const f = imageFixture(picture({ embed: embedded ? "image" : null, link: "remote" }) + shape({ id: "3" }));
    f.parts["deck/_rels/z.xml.rels"] = relationships([f.rel("layout", "slideLayout", "layout.xml"), f.rel("image", "image", "../media/raster.bin"), { ...f.rel("remote", "image", "https://example.invalid/pixel.png?x=1&y=2"), TargetMode: "External" }]);
    const s = await open(f);
    assert.equal(s.document.assets.length, embedded ? 1 : 0);
    assert.equal(s.document.slides[0].elements.length, embedded ? 2 : 1);
    const ref = s.report.inspection.package.relationships.find(r => r.id === "remote");
    assert.equal(ref.target, "https://example.invalid/pixel.png?x=1&y=2");
    assert.ok(findings(s).some(f => /not fetched or registered/.test(f.reason))); s.dispose();
  }
});

test("missing, wrong-type, wrong-mode and fragmented image relationships reject without publishing a session", async () => {
  for (const make of [
    f => f.rel("image", "slide", "a.xml"),
    f => f.rel("other", "image", "../media/raster.bin"),
    f => ({ ...f.rel("image", "image", "https://example.invalid/a.png"), TargetMode: "External" }),
    f => f.rel("image", "image", "../media/raster.bin#fragment"),
    f => f.rel("image", "image", "../media/missing.png"),
  ]) {
    const f = imageFixture();
    f.parts["deck/_rels/z.xml.rels"] = relationships([f.rel("layout", "slideLayout", "layout.xml"), make(f)]);
    await assert.rejects(open(f), { code: "import-invalid-package" });
  }
  const f = imageFixture(picture({ embed: null, link: "image" }));
  await assert.rejects(open(f), { code: "import-invalid-package" });
});

test("PNG, JPEG and GIF are classified by bytes and MIME rather than suffix; mismatches and other formats stay opaque", async () => {
  const samples = [
    { data: png, mime: "image/png", mapped: true },
    { data: gif, mime: "image/gif", mapped: true },
    // Minimal headers deliberately prove classification is not full decoding.
    { data: jpegHeader, mime: "image/jpeg", mapped: true },
    { data: png, mime: "image/jpeg", mapped: false },
    { data: Buffer.from([0, 1, 2]), mime: "image/png", mapped: false },
    { data: Buffer.from("RIFFxxxxWEBP"), mime: "image/webp", mapped: false },
    { data: Buffer.from('<svg xmlns="http://www.w3.org/2000/svg"/>'), mime: "image/svg+xml", mapped: false },
  ];
  for (const { data, mime, mapped } of samples) {
    const f = imageFixture(picture() + shape({ id: "3" })); f.parts["media/raster.bin"] = data; f.mediaTypes["media/raster.bin"] = mime;
    const input = encodeImages(f), s = await importPptx(input, { projection: "content" });
    assert.equal(s.document.assets.length, mapped ? 1 : 0); assert.equal(s.document.slides[0].elements.length, mapped ? 2 : 1);
    if (mapped) assert.equal(s.document.assets[0].source.value, data.toString("base64"));
    else assert.ok(findings(s).some(f => /MIME\/signature/.test(f.reason)));
    const prepared = await s.prepareSave({ mode: "no-op" }); assert.deepEqual((await s.save(prepared.plan)).bytes, new Uint8Array(input)); s.dispose();
  }
});

test("real EMF and WMF pictures remain source-only and exact no-op saves preserve them", async () => {
  for (const format of ["emf", "wmf"]) {
    const data = await readFile(new URL(`../../../tests/fixtures/metafile-libreoffice-asymmetric.${format}`, import.meta.url));
    for (const strict of [false, true]) {
      const f = imageFixture(picture() + picture({ id: 4 }) + shape({ id: "3" }), strict);
      f.parts["media/raster.bin"] = data; f.mediaTypes["media/raster.bin"] = `image/${format}`;
      const input = encodeImages(f), s = await importPptx(input, { projection: "content" });
      assert.equal(s.document.assets.length, 0);
      assert.deepEqual(s.document.slides[0].elements.map(e => e.type), ["shape"]);
      assert.equal(findings(s).filter(f => /MIME\/signature/.test(f.reason)).length, 2);
      const plan = await s.prepareSave({ mode: "no-op" });
      assert.equal(plan.status, "ready");
      assert.deepEqual((await s.save(plan.plan)).bytes, new Uint8Array(input));
      assert.equal((await s.prepareSave({ mode: "rebuild" })).status, "blocked");
      s.dispose();
    }
  }
});

test("SVG picture extension explicitly identifies the raster-only projection and preserves source on no-op", async () => {
  for (const strict of [false, true]) for (const external of [false, true]) {
    const extension = '<a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><v:svgBlip xmlns:v="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="svg"/></a:ext></a:extLst>';
    const f = imageFixture(picture({ effects: extension }), strict);
    const svg = Buffer.from('<svg xmlns="http://www.w3.org/2000/svg"><rect width="8" height="4"/></svg>');
    f.parts["media/vector.svg"] = svg; f.mediaTypes["media/vector.svg"] = "image/svg+xml";
    f.parts["deck/_rels/z.xml.rels"] = relationships([f.rel("layout", "slideLayout", "layout.xml"), f.rel("image", "image", "../media/raster.bin"), external ? { ...f.rel("svg", "image", "https://example.invalid/vector.svg"), TargetMode: "External" } : f.rel("svg", "image", "../media/vector.svg")]);
    const input = encodeImages(f), s = await importPptx(input, { projection: "content" });
    assert.equal(s.document.assets.length, 1);
    assert.equal(s.document.assets[0].source.value, png.toString("base64"));
    assert.ok(findings(s).some(f => /only a supported base image relationship can be projected/.test(f.reason)));
    assert.ok(findings(s).some(f => /Source subtree is not interpreted/.test(f.reason) && f.source.path?.includes("extLst")));
    const plan = await s.prepareSave({ mode: "no-op" }); assert.equal(plan.status, "ready");
    assert.deepEqual((await s.save(plan.plan)).bytes, new Uint8Array(input)); s.dispose();
  }
  const extension = '<a:extLst><a:ext uri="urn:test"><v:svgBlip xmlns:v="http://schemas.microsoft.com/office/drawing/2016/SVG/main"/></a:ext></a:extLst>';
  const missing = await open(imageFixture(picture({ embed: null, effects: extension })));
  assert.equal(missing.document.assets.length, 0);
  assert.ok(findings(missing).some(f => /only a supported base image relationship can be projected/.test(f.reason)));
  missing.dispose();
  const lookalike = await open(imageFixture(picture({ effects: extension.replace('http://schemas.microsoft.com/office/drawing/2016/SVG/main', 'urn:not-svg') })));
  assert.ok(!findings(lookalike).some(f => /An SVG image extension is present/.test(f.reason)));
  lookalike.dispose();
});

test("header dimensions are bounded and malformed raster headers remain source-only", async () => {
  const pngZero = Buffer.from(png); pngZero.writeUInt32BE(0, 16);
  const pngHuge = Buffer.from(png); pngHuge.writeUInt32BE(0x80000000, 20);
  const pngNoIhdr = Buffer.from(png); pngNoIhdr[12] = 0;
  const gifZero = Buffer.from(gif); gifZero.writeUInt16LE(0, 8);
  const badJpegLength = Buffer.from(jpegHeader); badJpegLength.writeUInt16BE(500, 4);
  const jpegZero = Buffer.from(jpegHeader); jpegZero.writeUInt16BE(0, 17);
  for (const [data, mime] of [[pngZero, "image/png"], [pngHuge, "image/png"], [pngNoIhdr, "image/png"], [png.subarray(0, 20), "image/png"], [gifZero, "image/gif"], [gif.subarray(0, 9), "image/gif"], [jpegHeader.subarray(0, 19), "image/jpeg"], [badJpegLength, "image/jpeg"], [jpegZero, "image/jpeg"], [Buffer.from([255, 216, 255, 218]), "image/jpeg"], [Buffer.from([255, 216, 255]), "image/jpeg"]]) {
    const f = imageFixture(); f.parts["media/raster.bin"] = data; f.mediaTypes["media/raster.bin"] = mime;
    const s = await open(f); assert.equal(s.document.assets.length, 0); assert.equal(s.document.slides[0].elements.length, 0);
    assert.ok(findings(s).some(f => /dimensions/.test(f.reason))); s.dispose();
  }
  const f = imageFixture(); f.parts["media/raster.bin"] = jpegHeader; f.mediaTypes["media/raster.bin"] = "image/jpeg";
  const s = await open(f); assert.equal(s.document.assets[0].width, 8); assert.equal(s.document.assets[0].height, 4); s.dispose();
});

test("unsupported placement, crop outsets and nonrectangular masks omit only their picture without orphan assets", async () => {
  for (const options of [
    { placement: "<a:tile/>" }, { placement: "" }, { preset: "ellipse" },
    { placement: '<a:stretch><a:fillRect l="10000"/></a:stretch>' },
    { crop: '<a:srcRect l="-10000"/>' }, { crop: '<a:srcRect l="50000" r="50000"/>' },
    { crop: '<a:srcRect t="110000"/>' }, { fillAttrs: 'rotWithShape="0"' },
  ]) {
    const s = await open(imageFixture(picture(options) + shape({ id: "3" })));
    assert.equal(s.document.assets.length, 0); assert.equal(s.document.slides[0].elements.length, 1);
    const f = findings(s).find(f => f.feature === "drawing.pic");
    assert.equal(f.tier, "unsupported"); assert.equal(f.source.sourceObjectId, "2"); s.dispose();
  }
});

test("malformed image percentages and conflicting placements fail; valid percent literals and empty crop defaults map", async () => {
  for (const options of [
    { crop: '<a:srcRect l="NaN"/>' }, { crop: '<a:srcRect l="1e4"/>' }, { crop: '<a:srcRect l="2147483648"/>' },
    { crop: '<a:srcRect l="10.1"/>' }, { crop: '<a:srcRect/><a:srcRect/>' },
    { effects: '<a:alphaModFix amt="-1"/>' }, { placement: '<a:stretch/><a:tile/>' }, { fillAttrs: 'rotWithShape="maybe"' },
  ]) await assert.rejects(open(imageFixture(picture(options))), { code: "import-invalid-package" });
  const s = await open(imageFixture(picture({ crop: '<a:srcRect l="12.5%" b="+20%"/>', effects: '<a:alphaModFix amt="50%"/>' }), true));
  assert.deepEqual(first(s).crop, { left: .125, top: 0, right: 0, bottom: .2 }); assert.equal(first(s).opacity, .5); s.dispose();
  const empty = await open(imageFixture(picture({ crop: "<a:srcRect/>", effects: "<a:alphaModFix/>" })));
  assert.equal(first(empty).opacity, 1); assert.deepEqual(first(empty).crop, { left: 0, top: 0, right: 0, bottom: 0 }); empty.dispose();
});

test("unknown image effects, locks and namespace-qualified attributes remain located without falsely consuming lookalikes", async () => {
  const f = imageFixture(picture({ effects: '<a:grayscl/><a:alphaModFix amt="200000"/>', props: '<a:effectLst><a:outerShdw/></a:effectLst>', blipAttrs: 'xmlns:q="urn:unknown" q:embed="not-an-image"' }));
  f.parts["deck/z.xml"] = f.parts["deck/z.xml"].replace("<p:cNvPicPr/>", '<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>');
  const s = await open(f); assert.equal(first(s).opacity, 1);
  for (const suffix of ["grayscl[1]", "effectLst[1]", "picLocks[1]", "@q:embed"]) assert.ok(findings(s).some(f => f.source.path?.endsWith(suffix) && f.elementId === first(s).id), suffix);
  assert.ok(findings(s).some(f => /amplification/.test(f.reason))); s.dispose();
});

test("image budgets count expanded copied bytes, retain old structure behavior and recover on a fresh import", async () => {
  const f = imageFixture();
  for (const key of ["maxImportedAssetBytes", "maxTotalImportedAssetBytes"]) {
    await assert.rejects(open(f, { sessionLimits: { [key]: png.length - 1 } }), { code: "import-limit-exceeded", source: { part: "/media/raster.bin" } });
    for (const value of [undefined, 0, -1, 1.5, Infinity]) await assert.rejects(open(f, { sessionLimits: { [key]: value } }), RangeError);
  }
  const s = await open(f, { sessionLimits: { maxImportedAssetBytes: png.length, maxTotalImportedAssetBytes: png.length } }); assert.equal(s.document.assets.length, 1); s.dispose();
  const structural = await open(f, { projection: "structure", sessionLimits: { maxImportedAssetBytes: 1 } });
  assert.equal(structural.document.assets.length, 0); assert.equal(structural.document.slides[0].elements.length, 0); structural.dispose();
  const controller = new AbortController(), pending = open(f, { signal: controller.signal }); controller.abort();
  await assert.rejects(pending, { code: "operation-aborted" });
});

test("unsupported MIME stays source-only without spending image projection budget", async () => {
  for (const strict of [false, true]) for (const unsupportedFirst of [false, true]) {
    const vector = picture({ id: 2, embed: "vector" }), raster = picture({ id: 3, embed: "image" });
    const f = imageFixture(unsupportedFirst ? vector + raster + vector.replace('id="2"', 'id="4"') : raster + vector, strict);
    f.parts["media/vector.svg"] = Buffer.from('<svg xmlns="http://www.w3.org/2000/svg">' + ' '.repeat(png.length * 3) + '</svg>');
    f.mediaTypes["media/vector.svg"] = "image/svg+xml";
    f.parts["deck/_rels/z.xml.rels"] = relationships([f.rel("layout", "slideLayout", "layout.xml"), f.rel("image", "image", "../media/raster.bin"), f.rel("vector", "image", "../media/vector.svg")]);
    const input = encodeImages(f), session = await importPptx(input, { projection: "content", sessionLimits: { maxImportedAssetBytes: png.length, maxTotalImportedAssetBytes: png.length } });
    assert.equal(session.document.assets.length, 1);
    assert.equal(session.document.assets[0].source.value, png.toString("base64"));
    assert.equal(session.document.slides[0].elements.length, 1);
    assert.equal(findings(session).filter(f => /MIME\/signature/.test(f.reason)).length, unsupportedFirst ? 2 : 1);
    const plan = await session.prepareSave({ mode: "no-op" }); assert.equal(plan.status, "ready");
    assert.deepEqual((await session.save(plan.plan)).bytes, new Uint8Array(input)); session.dispose();
    f.mediaTypes["media/vector.svg"] = "image/png";
    await assert.rejects(open(f, { sessionLimits: { maxImportedAssetBytes: png.length, maxTotalImportedAssetBytes: png.length } }), { code: "import-limit-exceeded" });
  }
});

test("asset replacement invalidates no-op plans; exact restoration and other sessions retain original bytes", async () => {
  const input = encodeImages(), original = new Uint8Array(input);
  const aPromise = importPptx(input, { projection: "content" }); input.fill(0);
  const [a, b] = await Promise.all([aPromise, importPptx(original, { projection: "content" })]);
  const savedAsset = a.document.assets[0], prepared = await a.prepareSave({ mode: "no-op" });
  a.document.replaceAsset(savedAsset.id, { kind: "image", mimeType: "image/gif", source: { type: "bytes", value: gif } });
  assert.equal((await a.prepareSave({ mode: "no-op" })).status, "blocked");
  await assert.rejects(a.save(prepared.plan), { code: "save-snapshot-stale" });
  assert.equal(b.document.assets[0].source.value, png.toString("base64"));
  const { id, ...old } = savedAsset; a.document.replaceAsset(id, old);
  assert.deepEqual((await a.save(prepared.plan)).bytes, original);
  assert.equal((await a.prepareSave({ mode: "edit" })).status, "blocked");
  a.dispose(); assert.equal(first(b).type, "image"); b.dispose();
});

test("image relationship namespace aliases and MCE-selected pictures resolve expanded names", async () => {
  const f = imageFixture();
  f.parts["deck/z.xml"] = f.parts["deck/z.xml"].replace(picture(), `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:q="urn:future"><mc:Choice Requires="q">${picture({ embed: "not-selected" })}</mc:Choice><mc:Fallback>${picture()}</mc:Fallback></mc:AlternateContent>`).replaceAll("xmlns:r=", "xmlns:rel=").replaceAll("r:embed=", "rel:embed=");
  const s = await open(f); assert.equal(first(s).type, "image");
  assert.equal(findings(s).filter(f => f.source.path?.endsWith("/@rel:embed")).length, 0); s.dispose();
});
