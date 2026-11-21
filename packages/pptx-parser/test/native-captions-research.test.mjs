import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile, writeFile, mkdtemp, rm } from 'node:fs/promises';
import { execFileSync } from 'node:child_process';
import { createRequire } from 'node:module';
import { createHash } from 'node:crypto';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { importPptx, inspectPptxCaptionTracks } from '../dist/index.js';

const root = fileURLToPath(new URL('../../../', import.meta.url));
const { zipSync, unzipSync, strFromU8, strToU8 } = createRequire(import.meta.url)('fflate');
const sample = name => readFile(join(root, `tests/fixtures/${name}.pptx`));
const inventory = file => JSON.parse(execFileSync('python3', [join(root, 'scripts/research/inspect-native-captions.py'), file], { encoding: 'utf8' }));

test('PowerPoint caption samples inventory zero, single and multiple tracks without inferring a default', async () => {
  const empty = inventory(join(root, 'tests/fixtures/pptkit-video-trim.pptx'));
  assert.deepEqual(empty.map(x => x.tracks.length), [0, 0]);
  const one = inventory(join(root, 'tests/fixtures/powerpoint-caption-single.pptx'));
  const multi = inventory(join(root, 'tests/fixtures/powerpoint-caption-multi.pptx'));
  assert.deepEqual(one.map(x => x.tracks.length), [1, 0]);
  assert.deepEqual(multi.map(x => x.tracks.length), [2, 0]);
  assert.deepEqual(multi[0].trims, [{ st: '150', end: '50' }]);
  assert.deepEqual(multi[0].trackInfo, [{ displayLoc: 'media' }]);
  assert.deepEqual(multi[0].tracks.map(({ language, label, status }) => [language, label, status]), [['en-us', '', 'vtt-header-only'], ['fr-fr', 'French test', 'vtt-header-only']]);
  assert.equal(multi[0].tracks[0].id, one[0].tracks[0].id);
  assert.match(multi[0].tracks[0].text, /00:00:00\.100 --> 00:00:00\.900/);
  assert.match(multi[0].tracks[1].text, /Première légende/);
  for (const name of ['powerpoint-caption-single', 'powerpoint-caption-multi']) {
    const bytes = await sample(name), session = await importPptx(bytes, { projection: 'content', videoProjection: 'iso-bmff-v1' });
    try {
      const findings = session.report.findings.filter(x => x.feature === 'source-part-content' && x.operation === 'import' && x.tier === 'unsupported');
      assert.equal(findings.filter(x => /track\d+\.vtt$/.test(x.source.part)).length, name.endsWith('single') ? 1 : 2);
      const plan = await session.prepareSave({ mode: 'no-op' }); assert.deepEqual((await session.save(plan.plan)).bytes, new Uint8Array(bytes));
    } finally { session.dispose(); }
  }
});

test('SRT inserted on a second shared video is converted to a separate WebVTT track', async () => {
  const before = inventory(join(root, 'tests/fixtures/powerpoint-caption-multi.pptx'));
  const after = inventory(join(root, 'tests/fixtures/powerpoint-caption-srt.pptx'));
  const source = await readFile(join(root, 'tests/fixtures/caption-spanish-source.srt'), 'utf8');
  assert.match(source, /00:00:00,100 --> 00:00:00,900/);
  assert.deepEqual(after.map(x => x.tracks.length), [2, 1]);
  assert.deepEqual(after[0].tracks.map(({ id, language, label, text }) => ({ id, language, label, text })), before[0].tracks.map(({ id, language, label, text }) => ({ id, language, label, text })));
  const track = after[1].tracks[0]; assert.equal(track.language, 'es-es'); assert.equal(track.contentType, 'text/vtt'); assert.equal(track.status, 'vtt-header-only');
  assert.match(track.text, /00:00:00\.100 --> 00:00:00\.900\nPrimera frase/);
  assert.match(track.text, /00:00:01\.000 --> 00:00:01\.800\nSegunda frase/);
  const bytes = await sample('powerpoint-caption-srt'), entries = unzipSync(bytes), original = unzipSync(await sample('powerpoint-caption-multi'));
  assert.equal(Object.keys(entries).filter(x => x.endsWith('.srt')).length, 0);
  const videos = Object.keys(entries).filter(x => x.endsWith('.mp4')); assert.equal(videos.length, 1); assert.deepEqual(entries[videos[0]], original[videos[0]]);
  const session = await importPptx(bytes, { projection: 'content', videoProjection: 'iso-bmff-v1' });
  try { const plan = await session.prepareSave({ mode: 'no-op' }); assert.deepEqual((await session.save(plan.plan)).bytes, new Uint8Array(bytes)); }
  finally { session.dispose(); }
});

test('native playback sample retains source-relative cue times and trim', () => {
  const [video] = inventory(join(root, 'tests/fixtures/powerpoint-caption-playback.pptx'));
  assert.deepEqual(video.trims, [{ st: '2000', end: '2000' }]);
  assert.match(video.tracks[0].text, /00:00:00\.000 --> 00:00:02\.000\nZERO TO TWO/);
  assert.match(video.tracks[0].text, /00:00:10\.000 --> 00:00:12\.000\nTEN TO TWELVE/);
});

test('caption playback observation is bound to the retained source and source-relative cues', async () => {
  const evidence = JSON.parse(await readFile(join(root, 'docs/testing/native-caption-preview-evidence.json'), 'utf8'));
  const bytes = await readFile(join(root, evidence.sourcePath)), track = await readFile(join(root, evidence.sourceTrackPath), 'utf8');
  assert.equal(createHash('sha256').update(bytes).digest('hex'), evidence.sourceSha256);
  const [video] = inventory(join(root, evidence.sourcePath));
  assert.deepEqual(video.trims, [{ st: String(evidence.trimStartMs), end: String(evidence.trimEndMs) }]);
  assert.equal(video.tracks.length, 1); assert.equal(video.tracks[0].language, evidence.sourceTrackLanguage); assert.equal(video.tracks[0].text, `${track}\n`);
  const timestamp = ms => `00:00:${String(Math.floor(ms / 1000)).padStart(2, '0')}.${String(ms % 1000).padStart(3, '0')}`;
  for (const observation of evidence.preview.observations) {
    const sourceTime = observation.displayedElapsedMs + evidence.trimStartMs;
    assert.ok(sourceTime >= observation.cueStartMs && sourceTime < observation.cueEndMs);
    assert.ok(track.includes(`${timestamp(observation.cueStartMs)} --> ${timestamp(observation.cueEndMs)}\n${observation.caption}`));
  }
  assert.equal(evidence.preview.initialMenuSelection, 'off'); assert.equal(evidence.preview.offSelectionHidText, true);
  assert.equal(evidence.slideshow.startSettingOnResearchCopy, 'automatic'); assert.ok(evidence.slideshow.observedCaptions.every(caption => track.includes(caption)));
  assert.match(evidence.scope, /screenshots were viewed but not retained/);
});

test('multitrack preview observations refer to retained per-instance captions and source cues', async () => {
  const evidence = JSON.parse(await readFile(join(root, 'docs/testing/native-caption-multitrack-evidence.json'), 'utf8'));
  const bytes = await readFile(join(root, evidence.sourcePath)); assert.equal(createHash('sha256').update(bytes).digest('hex'), evidence.sourceSha256);
  const inventoryResult = inventory(join(root, evidence.sourcePath)), entries = unzipSync(bytes);
  assert.equal(Object.keys(entries).filter(name => name.endsWith('.mp4')).length, 1);
  const slideXml = strFromU8(entries['ppt/slides/slide1.xml']), relXml = strFromU8(entries['ppt/slides/_rels/slide1.xml.rels']);
  assert.equal([...slideXml.matchAll(/<p14:media\b[^>]*r:embed="rId2"/gu)].length, 2);
  assert.match(relXml, /Id="rId2"[^>]*Target="\.\.\/media\/media1\.mp4"/u);
  const timestamp = ms => `00:00:${String(Math.floor(ms / 1000)).padStart(2, '0')}.${String(ms % 1000).padStart(3, '0')}`;
  for (const instance of evidence.instances) {
    const item = inventoryResult.find(result => result.objectId === instance.objectId); assert.ok(item);
    assert.deepEqual(item.tracks.map(track => track.language), instance.trackLanguages);
    assert.equal(Number(item.trims[0]?.st ?? 0), instance.trimStartMs);
    for (const observation of instance.observations) {
      const sourceTime = observation.displayedElapsedMs + instance.trimStartMs, track = item.tracks.find(value => value.language === observation.language); assert.ok(track);
      assert.ok(sourceTime >= observation.cueStartMs && sourceTime < observation.cueEndMs);
      assert.ok(track.text.includes(`${timestamp(observation.cueStartMs)} --> ${timestamp(observation.cueEndMs)}\n${observation.caption}`));
    }
  }
  assert.deepEqual(evidence.menuObservations.firstAvailable, ['off', 'en-us', 'fr-fr']);
  assert.deepEqual(evidence.menuObservations.secondAvailable, ['off', 'es-es']);
  assert.equal(inventoryResult[0].tracks[1].label, evidence.menuObservations.firstFrenchLabel);
  assert.match(evidence.scope, /screenshots were viewed but not retained/);
});

test('loop observation is bound to retained tracks, trim and native timing settings', async () => {
  const evidence = JSON.parse(await readFile(join(root, 'docs/testing/native-caption-loop-evidence.json'), 'utf8'));
  const bytes = await readFile(join(root, evidence.sourcePath)); assert.equal(createHash('sha256').update(bytes).digest('hex'), evidence.sourceSha256);
  const [video] = inventory(join(root, evidence.sourcePath)); assert.deepEqual(video.tracks.map(track => track.language), evidence.trackLanguages);
  assert.deepEqual(video.trims, [{ st: String(evidence.trimStartMs), end: String(evidence.trimEndMs) }]);
  const french = await readFile(join(root, evidence.frenchInputPath), 'utf8'); assert.equal(video.tracks[1].text, `${french}\n`);
  const slideXml = strFromU8(unzipSync(bytes)['ppt/slides/slide1.xml']); assert.match(slideXml, /repeatCount="indefinite"/u);
  const timestamp = ms => `00:00:${String(Math.floor(ms / 1000)).padStart(2, '0')}.${String(ms % 1000).padStart(3, '0')}`;
  for (const item of [evidence.preview.beforeRollover, evidence.preview.afterRollover]) {
    const sourceTime = item.elapsedMs + evidence.trimStartMs; assert.ok(sourceTime >= item.cueStartMs && sourceTime < item.cueEndMs);
    assert.ok(french.includes(`${timestamp(item.cueStartMs)} --> ${timestamp(item.cueEndMs)}\n${item.caption}`));
  }
  assert.ok(evidence.preview.afterRollover.elapsedMs < evidence.preview.beforeRollover.elapsedMs);
  assert.equal(evidence.slideshow.applicationExitedUnexpectedly, true); assert.equal(evidence.slideshow.crashCause, 'undetermined');
  assert.deepEqual((await inspectPptxCaptionTracks(bytes)).tracks.map(track => track.language), evidence.trackLanguages);
  const session = await importPptx(bytes, { projection: 'content' });
  try { const plan = await session.prepareSave({ mode: 'no-op' }); assert.deepEqual((await session.save(plan.plan)).bytes, new Uint8Array(bytes)); } finally { session.dispose(); }
});

test('research inventory identifies damaged track content and relationships without loading external targets', async () => {
  const original = await sample('powerpoint-caption-multi'), dir = await mkdtemp(join(tmpdir(), 'pptkit-caption-research-'));
  const cases = [
    [entries => { entries['ppt/media/track1.vtt'] = strToU8('not a WebVTT file'); }, 'invalid-vtt-header'],
    [entries => { entries['ppt/media/track1.vtt'] = Uint8Array.from([0xff, 0xfe]); }, 'invalid-utf8'],
    [entries => { delete entries['ppt/media/track1.vtt']; }, 'missing-or-invalid-target'],
    [entries => { const key = 'ppt/slides/_rels/slide1.xml.rels'; entries[key] = strToU8(strFromU8(entries[key]).replace('Target="../media/track1.vtt"', 'Target="https://example.invalid/captions.vtt" TargetMode="External"')); }, 'external-not-loaded'],
    [entries => { const key = 'ppt/slides/_rels/slide1.xml.rels'; entries[key] = strToU8(strFromU8(entries[key]).replace(/(Id="rId6" Type="[^"]*2017\/04\/relationships\/)track/, '$1other')); }, 'invalid-relationship'],
    [entries => { const key = '[Content_Types].xml'; entries[key] = strToU8(strFromU8(entries[key]).replace('ContentType="text/vtt"', 'ContentType="text/plain"')); }, 'unsupported-content-type'],
  ];
  try {
    for (const [mutate, status] of cases) {
      const entries = unzipSync(original); mutate(entries); const path = join(dir, `${status}.pptx`); await writeFile(path, zipSync(entries));
      const result = inventory(path); assert.equal(result[0].tracks[0].status, status); assert.equal(result[0].tracks[1].status, status === 'unsupported-content-type' ? status : 'vtt-header-only');
      if (status === 'invalid-utf8') {
        const bytes = await readFile(path), session = await importPptx(bytes, { projection: 'content', videoProjection: 'iso-bmff-v1' });
        try { const plan = await session.prepareSave({ mode: 'no-op' }); assert.deepEqual((await session.save(plan.plan)).bytes, new Uint8Array(bytes)); assert.ok(session.report.findings.some(x => x.feature === 'source-part-content' && x.operation === 'import' && x.tier === 'unsupported' && x.source.part.endsWith('/track1.vtt'))); } finally { session.dispose(); }
      }
    }
  } finally { await rm(dir, { recursive: true, force: true }); }
});
