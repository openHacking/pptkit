import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createRequire } from 'node:module';
import { importPptx } from '../dist/index.js';
import { importPptxFile } from '../dist/node.js';

const { unzipSync, zipSync, strFromU8, strToU8 } = createRequire(import.meta.url)('fflate');
const fixture = () => readFile(new URL('../../../tests/fixtures/h264-mp4-powerpoint.pptx', import.meta.url));
const options = { projection: 'content', videoSourceProfile: 'iso-bmff-v1' };
const slide = 'ppt/slides/slide1.xml', rels = 'ppt/slides/_rels/slide1.xml.rels';
function edit(bytes, part, change) {
  const parts = unzipSync(bytes); parts[part] = strToU8(change(strFromU8(parts[part]))); return zipSync(parts);
}
function change(text, old, replacement) { assert.ok(text.includes(old)); return text.replace(old, replacement); }

test('native video inventory locates separate video and poster sources without projecting playback', async () => {
  const bytes = await fixture(), s = await importPptx(bytes, options);
  try {
    const [video, media, poster] = s.report.videoSources;
    assert.deepEqual(s.report.videoSources.map(x => x.declaration), ['videoFile', 'p14:media', 'poster']);
    assert.equal(video.status, 'iso-bmff-framing'); assert.equal(video.mediaPart, '/ppt/media/media1.mp4');
    assert.equal(video.source.sourceObjectId, '4'); assert.match(video.source.path, /videoFile\[1\]$/);
    assert.deepEqual(video.framing, { majorBrand: 'isom', minorVersion: 512, compatibleBrands: ['isom', 'iso2', 'avc1', 'mp41'], boxCount: 4 });
    assert.ok(Object.isFrozen(video.framing.compatibleBrands));
    assert.equal(media.companionStatus, 'same-part'); assert.equal(media.mediaPart, video.mediaPart);
    assert.equal(poster.status, 'internal'); assert.equal(poster.mediaPart, '/ppt/media/image1.png'); assert.equal(poster.contentType, 'image/png');
    assert.equal(s.document.slides[0].elements.find(x => x.id === s.report.diagnostics.find(d => d.code === 'media-playback-unmapped').elementId).type, 'image');
    const plan = await s.prepareSave({ mode: 'no-op' }); assert.deepEqual((await s.save(plan.plan)).bytes, new Uint8Array(bytes));
  } finally { s.dispose(); }
  const normal = await importPptx(bytes, { projection: 'content' }); try { assert.equal(normal.report.videoSources, undefined); } finally { normal.dispose(); }
});

test('shared media is inventoried per object but charged once', async () => {
  const bytes = await fixture(), parts = unzipSync(bytes), xml = strFromU8(parts[slide]);
  const pic = xml.match(/<p:pic>.*?<\/p:pic>/u)[0], second = change(pic, 'id="4" name=', 'id="5" name=');
  parts[slide] = strToU8(change(xml, pic, pic + second));
  const input = zipSync(parts), mediaSize = parts['ppt/media/media1.mp4'].length;
  const s = await importPptx(input, { ...options, sessionLimits: { maxTotalImportedAssetBytes: mediaSize } });
  try { assert.equal(s.report.videoSources.length, 6); assert.deepEqual(s.report.videoSources.filter(x => x.declaration === 'videoFile').map(x => x.source.sourceObjectId), ['4', '5']); } finally { s.dispose(); }
  await assert.rejects(importPptx(input, { ...options, sessionLimits: { maxImportedAssetBytes: mediaSize - 1 } }), { code: 'import-limit-exceeded' });
  await assert.rejects(importPptx(input, { ...options, sessionLimits: { maxTotalImportedAssetBytes: mediaSize - 1 } }), { code: 'import-limit-exceeded' });
});

test('video inventory reaches an unprojectable group and does not misclassify audio', async () => {
  const bytes = await fixture();
  const grouped = edit(bytes, slide, xml => {
    const pic = xml.match(/<p:pic>.*?<\/p:pic>/u)[0];
    return change(xml, pic, `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="9" name="unmapped"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${pic}</p:grpSp>`);
  });
  const group = await importPptx(grouped, options);
  try { assert.equal(group.report.videoSources.length, 3); assert.equal(group.report.videoSources[0].source.sourceObjectId, '4'); } finally { group.dispose(); }
  const audioOnly = edit(bytes, slide, xml => change(xml, '<a:videoFile r:link="rId2"/>', '<a:audioFile r:link="rId2"/>'));
  const audio = await importPptx(audioOnly, options);
  try { assert.deepEqual(audio.report.videoSources, []); } finally { audio.dispose(); }
});

test('missing, external, mismatched and malformed video sources remain explicit', async () => {
  const original = await fixture();
  for (const mode of ['missing-video', 'external-video', 'external-poster', 'different-companion', 'conflicting-declaration', 'wrong-mime', 'damaged', 'no-poster']) {
    let bytes = original;
    if (mode === 'missing-video') bytes = edit(bytes, slide, xml => change(xml, '<a:videoFile r:link="rId2"', '<a:videoFile r:link="missing"'));
    if (mode === 'external-video') bytes = edit(bytes, rels, xml => change(xml, 'Target="../media/media1.mp4"/><Relationship Id="rId1"', 'Target="https://example.invalid/video.mp4" TargetMode="External"/><Relationship Id="rId1"'));
    if (mode === 'external-poster') {
      bytes = edit(bytes, slide, xml => change(xml, '<a:blip r:embed="rId4"', '<a:blip r:link="rId4"'));
      bytes = edit(bytes, rels, xml => change(xml, 'Target="../media/image1.png"', 'Target="https://example.invalid/poster.png" TargetMode="External"'));
    }
    if (mode === 'different-companion') {
      const parts = unzipSync(bytes); parts['ppt/media/other.mp4'] = parts['ppt/media/media1.mp4'];
      parts[rels] = strToU8(change(strFromU8(parts[rels]), 'Id="rId1" Type="http://schemas.microsoft.com/office/2007/relationships/media" Target="../media/media1.mp4"', 'Id="rId1" Type="http://schemas.microsoft.com/office/2007/relationships/media" Target="../media/other.mp4"')); bytes = zipSync(parts);
    }
    if (mode === 'wrong-mime') bytes = edit(bytes, '[Content_Types].xml', xml => change(xml, 'ContentType="video/mp4"', 'ContentType="application/octet-stream"'));
    if (mode === 'conflicting-declaration') bytes = edit(bytes, slide, xml => change(xml, '<a:videoFile r:link="rId2"/>', '<a:audioFile r:link="rId2"/><a:videoFile r:link="rId2"/>'));
    if (mode === 'damaged') { const parts = unzipSync(bytes); parts['ppt/media/media1.mp4'][4] = 0; bytes = zipSync(parts); }
    if (mode === 'no-poster') bytes = edit(bytes, slide, xml => change(xml, '<a:blip r:embed="rId4"/>', '<a:blip/>'));
    const s = await importPptx(bytes, options);
    try {
      const [video, companion, poster] = s.report.videoSources;
      if (mode === 'missing-video') assert.equal(video.status, 'unsupported');
      if (mode === 'external-video') { assert.equal(video.status, 'external'); assert.equal(video.externalTarget, 'https://example.invalid/video.mp4'); assert.equal(companion.companionStatus, 'unverified'); }
      if (mode === 'external-poster') { assert.equal(poster.status, 'external'); assert.equal(poster.externalTarget, 'https://example.invalid/poster.png'); }
      if (mode === 'different-companion') assert.equal(companion.companionStatus, 'different-part');
      if (mode === 'conflicting-declaration') assert.equal(companion.companionStatus, 'unverified');
      if (mode === 'wrong-mime' || mode === 'damaged') assert.equal(video.status, 'unsupported');
      if (mode === 'no-poster') assert.equal(poster.status, 'unsupported');
    } finally { s.dispose(); }
  }
  await assert.rejects(importPptx(original, { projection: 'structure', videoSourceProfile: 'iso-bmff-v1' }), TypeError);
  await assert.rejects(importPptx(original, { projection: 'content', videoSourceProfile: 'unknown' }), TypeError);
});

test('file adapter captures the video inventory option before asynchronous I/O', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'pptkit-video-'));
  try {
    const path = join(dir, 'sample.pptx'); await writeFile(path, await fixture());
    const config = { ...options }, pending = importPptxFile(path, config); config.videoSourceProfile = undefined;
    const { session } = await pending; try { assert.equal(session.report.videoSources.length, 3); } finally { session.dispose(); }
  } finally { await rm(dir, { recursive: true, force: true }); }
});
