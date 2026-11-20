import test from "node:test";
import assert from "node:assert/strict";
import { importPptx } from "../dist/index.js";
import { picture, imageFixture, encodeImages } from "./image-fixture.mjs";
import { relationships } from "./opc-fixture.mjs";
import { importPptxFile } from "../dist/node.js";
import { mkdtemp, readFile, writeFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { generatePptx } from '../../pptx-exporter/dist/index.js';
import { readGeneratedZip } from '../../../tests/support/read-generated-zip.mjs';

function fixture(strict = false) {
  const media = id => picture({ id }).replace('<p:nvPr/>', '<p:nvPr><a:audioFile r:link="sound"/></p:nvPr>');
  const f = imageFixture(media(2) + media(3), strict), wave = Buffer.alloc(124, 128);
  wave.write('RIFF', 0); wave.writeUInt32LE(116, 4); wave.write('WAVEfmt ', 8); wave.writeUInt32LE(16, 16);
  wave.writeUInt16LE(1, 20); wave.writeUInt16LE(1, 22); wave.writeUInt32LE(8000, 24); wave.writeUInt32LE(8000, 28);
  wave.writeUInt16LE(1, 32); wave.writeUInt16LE(8, 34); wave.write('data', 36); wave.writeUInt32LE(80, 40);
  f.parts['media/tone.wav'] = wave; f.mediaTypes['media/tone.wav'] = 'audio/wav';
  f.parts['deck/_rels/z.xml.rels'] = relationships([f.rel('layout', 'slideLayout', 'layout.xml'), f.rel('image', 'image', '../media/raster.bin'), f.rel('sound', 'audio', '../media/tone.wav')]);
  return f;
}
const options = { projection: 'content', audioSourceProfile: 'wave-pcm-v1' };
const projected = { projection: 'content', audioProjection: 'wave-pcm-v1' };

test('native PCM audio projects to detached Core audio and fresh output retains its bytes', async () => {
  const bytes = await readFile(new URL('../../../tests/fixtures/pcm-wave-powerpoint.pptx', import.meta.url));
  const session = await importPptx(bytes, projected);
  try {
    const element = session.document.slides[0].elements.find(x => x.type === 'audio');
    assert.ok(element); assert.equal(session.document.assets.find(x => x.id === element.audioAssetId).kind, 'audio');
    assert.equal(session.document.assets.find(x => x.id === element.posterAssetId).kind, 'image');
    assert.equal(session.report.findings.find(x => x.feature === 'drawing.audio.projected-fields' && x.operation === 'import').tier, 'fallback');
    assert.equal(session.report.findings.find(x => x.feature === 'drawing.audio.projected-fields' && x.operation === 'edit-save').tier, 'unsupported');
    assert.match(session.report.diagnostics.find(x => x.code === 'media-playback-unmapped').message, /playback settings and timing/);
    const plan = await session.prepareSave({ mode: 'no-op' }); assert.deepEqual((await session.save(plan.plan)).bytes, new Uint8Array(bytes));
    assert.equal((await session.prepareSave({ mode: 'edit' })).status, 'blocked');
    const fresh = await generatePptx(session.document, { audioMode: 'wave-pcm-v1' });
    const wave = readGeneratedZip(bytes).get('ppt/media/media1.wav'), output = readGeneratedZip(fresh.bytes);
    assert.ok([...output].some(([name, data]) => name.endsWith('.wav') && data.equals(wave)));
    let roundBytes = fresh.bytes;
    for (let round = 0; round < 3; round++) {
      const reopened = await importPptx(roundBytes, projected);
      try {
        assert.equal(reopened.document.slides[0].elements.filter(x => x.type === 'audio').length, 1);
        assert.equal(reopened.report.audioSources.filter(x => x.companionStatus === 'same-part').length, 1);
        const save = await reopened.prepareSave({ mode: 'no-op' }); assert.deepEqual((await reopened.save(save.plan)).bytes, roundBytes);
        roundBytes = (await generatePptx(reopened.document, { audioMode: 'wave-pcm-v1' })).bytes;
        assert.ok([...readGeneratedZip(roundBytes)].some(([name, data]) => name.endsWith('.wav') && data.equals(wave)));
      } finally { reopened.dispose(); }
    }
    const changed = structuredClone(element); changed.box.x += 1; const slide = session.document.slides[0];
    const index = slide.elements.findIndex(x => x.id === element.id); slide.removeElement(element.id); slide.insertElement(index, changed);
    assert.equal((await session.prepareSave({ mode: 'no-op' })).status, 'blocked');
  } finally { session.dispose(); }
});

test('shared embedded audio projects once with separate posters and a combined byte budget', async () => {
  for (const strict of [false, true]) {
    const f = fixture(strict), bytes = encodeImages(f), session = await importPptx(bytes, projected);
    try {
      const elements = session.document.slides[0].elements; assert.deepEqual(elements.map(x => x.type), ['audio', 'audio']);
      assert.equal(elements[0].audioAssetId, elements[1].audioAssetId); assert.equal(elements[0].posterAssetId, elements[1].posterAssetId);
      assert.equal(session.document.assets.filter(x => x.kind === 'audio').length, 1); assert.equal(session.document.assets.filter(x => x.kind === 'image').length, 1);
      const total = f.parts['media/tone.wav'].length + f.parts['media/raster.bin'].length;
      const ok = await importPptx(bytes, { ...projected, sessionLimits: { maxTotalImportedAssetBytes: total } }); ok.dispose();
      await assert.rejects(importPptx(bytes, { ...projected, sessionLimits: { maxTotalImportedAssetBytes: total - 1 } }), { code: 'import-limit-exceeded' });
    } finally { session.dispose(); }
  }
});

test('unsupported, conflicting or unmapped sources do not fabricate Core audio', async () => {
  for (const mode of ['external', 'damaged', 'conflict', 'different-companion', 'no-geometry', 'no-poster']) {
    const f = fixture();
    if (mode === 'external') f.parts['deck/_rels/z.xml.rels'] = relationships([f.rel('layout', 'slideLayout', 'layout.xml'), f.rel('image', 'image', '../media/raster.bin'), { ...f.rel('sound', 'audio', 'https://example.invalid/tone.wav'), TargetMode: 'External' }]);
    if (mode === 'damaged') f.parts['media/tone.wav'][0] = 0;
    if (mode === 'conflict') f.parts['deck/z.xml'] = f.parts['deck/z.xml'].replaceAll('<a:audioFile r:link="sound"/>', '<a:audioFile r:link="sound"/><a:videoFile r:link="sound"/>');
    if (mode === 'different-companion') {
      f.parts['media/other.wav'] = Buffer.from(f.parts['media/tone.wav']); f.mediaTypes['media/other.wav'] = 'audio/wav';
      f.parts['deck/z.xml'] = f.parts['deck/z.xml'].replaceAll('<a:audioFile r:link="sound"/>', '<a:audioFile r:link="sound"/><p:extLst><p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}"><m:media xmlns:m="http://schemas.microsoft.com/office/powerpoint/2010/main" r:embed="office"/></p:ext></p:extLst>');
      f.parts['deck/_rels/z.xml.rels'] = relationships([f.rel('layout', 'slideLayout', 'layout.xml'), f.rel('image', 'image', '../media/raster.bin'), f.rel('sound', 'audio', '../media/tone.wav'), { Id: 'office', Type: 'http://schemas.microsoft.com/office/2007/relationships/media', Target: '../media/other.wav' }]);
    }
    if (mode === 'no-geometry') f.parts['deck/z.xml'] = f.parts['deck/z.xml'].replaceAll(/<a:xfrm[^>]*>.*?<\/a:xfrm>/g, '');
    if (mode === 'no-poster') f.parts['deck/z.xml'] = f.parts['deck/z.xml'].replaceAll('r:embed="image"', '');
    const session = await importPptx(encodeImages(f), projected);
    try {
      const audio = session.document.slides[0].elements.filter(x => x.type === 'audio');
      assert.equal(audio.length, mode === 'no-poster' ? 2 : 0, mode);
      if (mode === 'no-poster') assert.ok(audio.every(x => x.posterAssetId === undefined));
    } finally { session.dispose(); }
  }
  await assert.rejects(importPptx(encodeImages(fixture()), { ...projected, editSaveProfile: 'text-runs-v1' }), TypeError);
  await assert.rejects(importPptx(encodeImages(fixture()), { ...projected, pictureSourceProfile: 'embedded-v1' }), TypeError);
  await assert.rejects(importPptx(encodeImages(fixture()), { ...projected, projection: 'structure' }), TypeError);
});

test('PowerPoint-created PCM WAVE retains its native paired relationships and poster', async () => {
  const bytes = await readFile(new URL('../../../tests/fixtures/pcm-wave-powerpoint.pptx', import.meta.url));
  const session = await importPptx(bytes, options);
  try {
    const [audio, media] = session.report.audioSources;
    assert.equal(session.report.audioSources.length, 2);
    assert.equal(audio.declaration, 'audioFile'); assert.equal(audio.status, 'wave-pcm');
    assert.equal(media.declaration, 'p14:media'); assert.equal(media.companionStatus, 'same-part');
    assert.equal(audio.mediaPart, '/ppt/media/media1.wav'); assert.equal(audio.contentType, 'audio/x-wav');
    assert.deepEqual(audio.pcm, { channels: 1, sampleRate: 8000, bitsPerSample: 16, frameCount: 8000, durationSeconds: 1 });
    const diagnostic = session.report.diagnostics.find(x => x.code === 'media-playback-unmapped');
    assert.equal(diagnostic.source.sourceObjectId, audio.source.sourceObjectId);
    assert.equal(session.document.slides[0].elements.find(x => x.id === diagnostic.elementId).type, 'image');
    const plan = await session.prepareSave({ mode: 'no-op' });
    assert.deepEqual((await session.save(plan.plan)).bytes, new Uint8Array(bytes));
  } finally { session.dispose(); }
});

test('audio source inventory identifies shared PCM and preserves original bytes', async () => {
  for (const strict of [false, true]) {
    const bytes = encodeImages(fixture(strict)), session = await importPptx(bytes, options);
    try {
      assert.equal(session.report.audioSources.length, 2);
      assert.equal(session.document.assets.length, 1); // Only the picture poster is projected.
      for (const [index, item] of session.report.audioSources.entries()) {
        assert.equal(item.status, 'wave-pcm'); assert.equal(item.mediaPart, '/media/tone.wav');
        assert.equal(item.relationshipId, 'sound'); assert.equal(item.source.sourceObjectId, String(index + 2));
        assert.match(item.source.path, /audioFile\[1\]$/); assert.ok(item.sourceElementIndex > 0);
        assert.deepEqual(item.pcm, { channels: 1, sampleRate: 8000, bitsPerSample: 8, frameCount: 80, durationSeconds: .01 });
        assert.ok(Object.isFrozen(item.pcm));
      }
      const plan = await session.prepareSave({ mode: 'no-op' });
      assert.deepEqual((await session.save(plan.plan)).bytes, new Uint8Array(bytes));
    } finally { session.dispose(); }
  }
});

test('file adapter captures the opt-in source profile before I/O', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'pptkit-audio-'));
  try {
    const path = join(dir, 'source.pptx'); await writeFile(path, encodeImages(fixture()));
    const config = { ...options }, pending = importPptxFile(path, config); config.audioSourceProfile = undefined;
    const { session } = await pending;
    try { assert.equal(session.report.audioSources.length, 2); } finally { session.dispose(); }
    const projection = { ...projected }, pendingProjection = importPptxFile(path, projection); projection.audioProjection = undefined;
    const imported = await pendingProjection;
    try { assert.equal(imported.session.document.slides[0].elements[0].type, 'audio'); } finally { imported.session.dispose(); }
  } finally { await rm(dir, { recursive: true, force: true }); }
});

test('inventory reaches pictures inside unsupported groups without fabricating geometry', async () => {
  for (const strict of [false, true]) {
    const f = fixture(strict), xml = f.parts['deck/z.xml'];
    const first = picture({ id: 2 }).replace('<p:nvPr/>', '<p:nvPr><a:audioFile r:link="sound"/></p:nvPr>');
    const group = `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="9" name="Unmapped group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${first}</p:grpSp>`;
    f.parts['deck/z.xml'] = xml.replace(first, group);
    const bytes = encodeImages(f), session = await importPptx(bytes, options);
    try {
      assert.deepEqual(session.report.audioSources.map(x => x.source.sourceObjectId), ['2', '3']);
      assert.equal(session.report.audioSources[0].status, 'wave-pcm');
      assert.equal(session.document.slides[0].elements.length, 1);
      const plan = await session.prepareSave({ mode: 'no-op' }); assert.deepEqual((await session.save(plan.plan)).bytes, new Uint8Array(bytes));
    } finally { session.dispose(); }
  }
});

test('WAVE and Office media declarations are independently inventoried with relationship modes', async () => {
  const guid = '{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}';
  for (const strict of [false, true]) for (const mode of ['embedded', 'different-part', 'external', 'wrong-mode', 'wrong-type', 'wrong-uri']) {
    const f = fixture(strict);
    const ext = `<p:extLst><p:ext uri="${mode === 'wrong-uri' ? '{00000000-0000-0000-0000-000000000000}' : guid}"><m:media xmlns:m="http://schemas.microsoft.com/office/powerpoint/2010/main" r:${mode === 'external' ? 'link' : 'embed'}="office"/></p:ext></p:extLst>`;
    f.parts['deck/z.xml'] = f.parts['deck/z.xml'].replace('<a:audioFile r:link="sound"/>', `<a:wavAudioFile r:embed="sound"/>${ext}`);
    if (mode === 'different-part') { f.parts['media/other.wav'] = Buffer.from(f.parts['media/tone.wav']); f.mediaTypes['media/other.wav'] = 'audio/wav'; }
    const rel = { Id: 'office', Type: mode === 'wrong-type' ? 'urn:wrong' : 'http://schemas.microsoft.com/office/2007/relationships/media', Target: mode === 'external' || mode === 'wrong-mode' ? 'https://example.invalid/audio.wav' : mode === 'different-part' ? '../media/other.wav' : '../media/tone.wav', ...(mode === 'external' || mode === 'wrong-mode' ? { TargetMode: 'External' } : {}) };
    f.parts['deck/_rels/z.xml.rels'] = relationships([f.rel('layout', 'slideLayout', 'layout.xml'), f.rel('image', 'image', '../media/raster.bin'), f.rel('sound', 'audio', '../media/tone.wav'), rel]);
    const session = await importPptx(encodeImages(f), options);
    try {
      const sources = session.report.audioSources;
      assert.equal(sources[0].declaration, 'wavAudioFile'); assert.equal(sources[0].relationshipAttribute, 'embed'); assert.equal(sources[0].status, 'wave-pcm');
      if (mode === 'wrong-uri') assert.equal(sources.length, 2);
      else {
        assert.equal(sources.length, 3); const item = sources[1];
        assert.equal(item.declaration, 'p14:media'); assert.equal(item.relationshipAttribute, mode === 'external' ? 'link' : 'embed');
        assert.equal(item.status, mode === 'embedded' || mode === 'different-part' ? 'wave-pcm' : mode === 'external' ? 'external' : 'unsupported');
        assert.equal(item.companionStatus, mode === 'embedded' ? 'same-part' : mode === 'different-part' ? 'different-part' : 'unverified');
        assert.match(item.source.path, /media\[1\]$/);
      }
    } finally { session.dispose(); }
  }
});

test('video-only Office media and inactive audio choices do not become audio sources', async () => {
  const guid = '{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}';
  const f = fixture();
  const ext = `<p:extLst><p:ext uri="${guid}"><m:media xmlns:m="http://schemas.microsoft.com/office/powerpoint/2010/main" r:embed="office"/></p:ext></p:extLst>`;
  f.parts['deck/z.xml'] = f.parts['deck/z.xml']
    .replace('<a:audioFile r:link="sound"/>', `<a:videoFile r:link="sound"/>${ext}`)
    .replace('<a:audioFile r:link="sound"/>', '<mc:AlternateContent><mc:Choice Requires="x"><a:audioFile r:link="sound"/></mc:Choice><mc:Fallback/></mc:AlternateContent>')
    .replace('<p:sld ', '<p:sld xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future" mc:Ignorable="x" ');
  const session = await importPptx(encodeImages(f), options);
  try { assert.deepEqual(session.report.audioSources, []); } finally { session.dispose(); }
});

test('audio inventory reports external and unsupported audio without fetching or guessing', async () => {
  for (const mode of ['external', 'missing', 'type', 'damaged', 'mime']) {
    const f = fixture();
    if (mode === 'external') f.parts['deck/_rels/z.xml.rels'] = relationships([f.rel('layout', 'slideLayout', 'layout.xml'), f.rel('image', 'image', '../media/raster.bin'), { ...f.rel('sound', 'audio', 'https://example.invalid/tone.wav'), TargetMode: 'External' }]);
    if (mode === 'missing') f.parts['deck/z.xml'] = f.parts['deck/z.xml'].replaceAll('r:link="sound"', 'r:link="unknown"');
    if (mode === 'type') f.parts['deck/_rels/z.xml.rels'] = relationships([f.rel('layout', 'slideLayout', 'layout.xml'), f.rel('image', 'image', '../media/raster.bin'), f.rel('sound', 'image', '../media/tone.wav')]);
    if (mode === 'damaged') f.parts['media/tone.wav'][0] = 0;
    if (mode === 'mime') f.mediaTypes['media/tone.wav'] = 'audio/mpeg';
    const session = await importPptx(encodeImages(f), options);
    try { assert.deepEqual(session.report.audioSources.map(x => x.status), [mode === 'external' ? 'external' : 'unsupported', mode === 'external' ? 'external' : 'unsupported']); }
    finally { session.dispose(); }
  }
  await assert.rejects(importPptx(encodeImages(fixture()), { ...options, sessionLimits: { maxImportedAssetBytes: 123 } }), { code: 'import-limit-exceeded' });
  const shared = await importPptx(encodeImages(fixture()), { ...options, sessionLimits: { maxImportedAssetBytes: 124, maxTotalImportedAssetBytes: 124 } });
  assert.equal(shared.report.audioSources.length, 2); shared.dispose();
  const ordinary = await importPptx(encodeImages(fixture()), { projection: 'content' });
  assert.equal(ordinary.report.audioSources, undefined); ordinary.dispose();
  await assert.rejects(importPptx(encodeImages(fixture()), { ...options, audioSourceProfile: 'other' }), TypeError);
  await assert.rejects(importPptx(encodeImages(fixture()), { ...options, projection: 'structure' }), TypeError);
});
