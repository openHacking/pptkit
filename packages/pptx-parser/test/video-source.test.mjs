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
const projected = { projection: 'content', videoProjection: 'iso-bmff-v1' };
const slide = 'ppt/slides/slide1.xml', rels = 'ppt/slides/_rels/slide1.xml.rels';
function edit(bytes, part, change) {
  const parts = unzipSync(bytes); parts[part] = strToU8(change(strFromU8(parts[part]))); return zipSync(parts);
}
function change(text, old, replacement) { assert.ok(text.includes(old)); return text.replace(old, replacement); }

test('native video projects to detached video and poster assets without enabling source edits', async () => {
  const bytes = await fixture(), parts = unzipSync(bytes), s = await importPptx(bytes, projected);
  try {
    const element = s.document.slides[0].elements.find(x => x.type === 'video'); assert.ok(element);
    const video = s.document.assets.find(x => x.id === element.videoAssetId), poster = s.document.assets.find(x => x.id === element.posterAssetId);
    assert.equal(video.kind, 'video'); assert.equal(poster.kind, 'image');
    assert.deepEqual(Buffer.from(video.source.value, 'base64'), Buffer.from(parts['ppt/media/media1.mp4']));
    assert.equal(s.report.videoSources.filter(x => x.companionStatus === 'same-part').length, 1);
    assert.equal(s.report.findings.find(x => x.feature === 'drawing.video.projected-fields' && x.operation === 'edit-save').tier, 'unsupported');
    assert.match(s.report.diagnostics.find(x => x.code === 'media-playback-unmapped').message, /video and optional poster are projected/);
    const plan = await s.prepareSave({ mode: 'no-op' }); assert.deepEqual((await s.save(plan.plan)).bytes, new Uint8Array(bytes));
    assert.equal((await s.prepareSave({ mode: 'edit' })).status, 'blocked');
    const total = parts['ppt/media/media1.mp4'].length + parts['ppt/media/image1.png'].length;
    const exact = await importPptx(bytes, { ...projected, sessionLimits: { maxTotalImportedAssetBytes: total } }); exact.dispose();
    await assert.rejects(importPptx(bytes, { ...projected, sessionLimits: { maxTotalImportedAssetBytes: total - 1 } }), { code: 'import-limit-exceeded' });
    s.document.replaceAsset(video.id, { kind: 'video', mimeType: 'video/mp4', source: { type: 'bytes', value: Uint8Array.of(1, 2, 3) } });
    assert.equal((await s.prepareSave({ mode: 'no-op' })).status, 'blocked');
  } finally { s.dispose(); }
  await assert.rejects(importPptx(bytes, { ...projected, editSaveProfile: 'text-runs-v1' }), TypeError);
  await assert.rejects(importPptx(bytes, { ...projected, pictureSourceProfile: 'embedded-v1' }), TypeError);
});

test('native video inventory locates separate video and poster sources without projecting playback', async () => {
  const bytes = await fixture(), s = await importPptx(bytes, options);
  try {
    const [video, media, poster] = s.report.videoSources;
    assert.deepEqual(s.report.videoSources.map(x => x.declaration), ['videoFile', 'p14:media', 'poster']);
    assert.equal(video.status, 'iso-bmff-framing'); assert.equal(video.mediaPart, '/ppt/media/media1.mp4');
    assert.equal(video.movieDurationMs, 2000); assert.equal(media.movieDurationMs, 2000); assert.equal(poster.movieDurationMs, undefined);
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
  const projectedSession = await importPptx(input, projected);
  try {
    const videos = projectedSession.document.slides[0].elements.filter(x => x.type === 'video');
    assert.equal(videos.length, 2); assert.equal(videos[0].videoAssetId, videos[1].videoAssetId);
    assert.equal(projectedSession.document.assets.filter(x => x.kind === 'video').length, 1);
  } finally { projectedSession.dispose(); }
});

test('simultaneous audio and video projection shares one image and aggregate budget', async () => {
  const parts = unzipSync(await fixture()), xml = strFromU8(parts[slide]);
  const pic = xml.match(/<p:pic>.*?<\/p:pic>/u)[0];
  const audio = pic.replace('id="4" name=', 'id="5" name=').replace('a:videoFile r:link="rId2"', 'a:audioFile r:link="rId8"').replace('r:embed="rId1"', 'r:embed="rId9"');
  parts[slide] = strToU8(change(xml, pic, pic + audio));
  const relXml = strFromU8(parts[rels]);
  parts[rels] = strToU8(change(relXml, '</Relationships>', '<Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio" Target="../media/tone.wav"/><Relationship Id="rId9" Type="http://schemas.microsoft.com/office/2007/relationships/media" Target="../media/tone.wav"/></Relationships>'));
  parts['[Content_Types].xml'] = strToU8(change(strFromU8(parts['[Content_Types].xml']), '</Types>', '<Default Extension="wav" ContentType="audio/wav"/></Types>'));
  const wav = Buffer.alloc(124, 128); wav.write('RIFF'); wav.writeUInt32LE(116, 4); wav.write('WAVEfmt ', 8); wav.writeUInt32LE(16, 16); wav.writeUInt16LE(1, 20); wav.writeUInt16LE(1, 22); wav.writeUInt32LE(8000, 24); wav.writeUInt32LE(8000, 28); wav.writeUInt16LE(1, 32); wav.writeUInt16LE(8, 34); wav.write('data', 36); wav.writeUInt32LE(80, 40);
  parts['ppt/media/tone.wav'] = wav;
  const bytes = zipSync(parts), total = parts['ppt/media/media1.mp4'].length + parts['ppt/media/image1.png'].length + wav.length;
  const opts = { ...projected, audioProjection: 'wave-pcm-v1' }, s = await importPptx(bytes, { ...opts, sessionLimits: { maxTotalImportedAssetBytes: total } });
  try {
    const media = s.document.slides[0].elements.filter(x => x.type === 'video' || x.type === 'audio');
    assert.deepEqual(media.map(x => x.type), ['video', 'audio']); assert.equal(media[0].posterAssetId, media[1].posterAssetId);
  } finally { s.dispose(); }
  await assert.rejects(importPptx(bytes, { ...opts, sessionLimits: { maxTotalImportedAssetBytes: total - 1 } }), { code: 'import-limit-exceeded' });
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
    const projectedSession = await importPptx(bytes, projected);
    try {
      const videos = projectedSession.document.slides[0].elements.filter(x => x.type === 'video');
      const supported = mode === 'no-poster' || mode === 'external-poster';
      assert.equal(videos.length, supported ? 1 : 0, mode);
      if (supported) assert.equal(videos[0].posterAssetId, undefined);
    } finally { projectedSession.dispose(); }
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
    const projection = { ...projected }, pendingProjection = importPptxFile(path, projection); projection.videoProjection = undefined;
    const imported = await pendingProjection;
    try { assert.ok(imported.session.document.slides[0].elements.some(x => x.type === 'video')); } finally { imported.session.dispose(); }
  } finally { await rm(dir, { recursive: true, force: true }); }
});


test('unsupported movie clock leaves framing and projection intact', async () => {
  const parts = unzipSync(await fixture()), video = parts['ppt/media/media1.mp4'];
  const tag = Buffer.from(video).indexOf(Buffer.from('mvhd')); assert.ok(tag > 0);
  video[tag + 4] = 1;
  const s = await importPptx(zipSync(parts), projected);
  try {
    const source = s.report.videoSources.find(x => x.declaration === 'videoFile');
    assert.equal(source.status, 'iso-bmff-framing'); assert.equal(source.movieDurationMs, undefined);
    assert.ok(s.document.slides[0].elements.some(x => x.type === 'video'));
  } finally { s.dispose(); }
});
