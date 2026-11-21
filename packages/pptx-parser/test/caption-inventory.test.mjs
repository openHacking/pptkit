import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import { inspectPptxCaptionTracks, importPptx } from '../dist/index.js';

const { zipSync, unzipSync, strFromU8, strToU8 } = createRequire(import.meta.url)('fflate');
const source = () => readFile(new URL('../../../tests/fixtures/powerpoint-caption-srt.pptx', import.meta.url));
const mutate = async edit => { const entries = unzipSync(await source()); edit(entries); return zipSync(entries); };
const rels = 'ppt/slides/_rels/slide1.xml.rels';

test('caption source inventory identifies the observed per-video tracks without semantic projection', async () => {
  const bytes = await source(), result = await inspectPptxCaptionTracks(bytes);
  assert.equal(result.schemaVersion, 'pptkit-caption-source-v1'); assert.equal(result.dialect, 'transitional'); assert.equal(result.view, 'raw-only');
  assert.deepEqual(result.tracks.map(x => [x.objectId, x.language, x.label, x.status]), [['2', 'en-us', '', 'vtt-header-only'], ['2', 'fr-fr', 'French test', 'vtt-header-only'], ['3', 'es-es', '', 'vtt-header-only']]);
  assert.equal(result.inspectedBytes, result.tracks.reduce((sum, x) => sum + x.byteLength, 0));
  assert.ok(result.tracks.every(x => x.source.part === '/ppt/slides/slide1.xml' && x.source.elementIndex > 0 && x.source.path.includes('track[')));
  assert.ok(Object.isFrozen(result) && Object.isFrozen(result.tracks) && Object.isFrozen(result.tracks[0].source));
  const empty = await readFile(new URL('../../../tests/fixtures/pptkit-video-trim.pptx', import.meta.url)); assert.deepEqual((await inspectPptxCaptionTracks(empty)).tracks, []);
  const imported = await importPptx(bytes, { projection: 'content' }); try { const plan = await imported.prepareSave({ mode: 'no-op' }); assert.deepEqual((await imported.save(plan.plan)).bytes, new Uint8Array(bytes)); } finally { imported.dispose(); }
});

test('caption inventory reports malformed tracks, wrong relationship and external targets', async () => {
  const cases = [
    [entries => { entries['ppt/media/track1.vtt'] = strToU8('bad'); }, 'invalid-vtt-header'],
    [entries => { entries['ppt/media/track1.vtt'] = Uint8Array.from([0xff, 0xfe]); }, 'invalid-utf8'],
    [entries => { entries['[Content_Types].xml'] = strToU8(strFromU8(entries['[Content_Types].xml']).replace('ContentType="text/vtt"', 'ContentType="text/plain"')); }, 'unsupported-content-type'],
    [entries => { entries[rels] = strToU8(strFromU8(entries[rels]).replace(/(Id="rId6" Type="[^"]*2017\/04\/relationships\/)track/, '$1other')); }, 'unsupported-relationship'],
    [entries => { entries[rels] = strToU8(strFromU8(entries[rels]).replace('Target="../media/track1.vtt"', 'Target="https://example.invalid/captions.vtt" TargetMode="External"')); }, 'external'],
  ];
  for (const [edit, expected] of cases) {
    const result = await inspectPptxCaptionTracks(await mutate(edit)); assert.equal(result.tracks[0].status, expected);
    if (expected === 'external') assert.equal(result.tracks[0].externalTarget, 'https://example.invalid/captions.vtt');
  }
  await assert.rejects(inspectPptxCaptionTracks(await mutate(entries => { delete entries['ppt/media/track1.vtt']; })), error => error.code === 'package-invalid');
});

test('caption inventory applies work, record, text and deduplicated byte budgets', async () => {
  const bytes = await source(), expected = await inspectPptxCaptionTracks(bytes);
  for (const captionLimits of [{ maxTracks: 2 }, { maxWorkUnits: 2 }, { maxTextCharacters: 10 }, { maxTrackBytes: 10 }, { maxTotalTrackBytes: 10 }]) await assert.rejects(inspectPptxCaptionTracks(bytes, { captionLimits }), error => error.code === 'package-limit-exceeded');
  await assert.rejects(inspectPptxCaptionTracks(bytes, { captionLimits: { maxTracks: 0 } }), RangeError);
  await assert.rejects(inspectPptxCaptionTracks(bytes, { captionLimits: { unknown: 1 } }), TypeError);
  const shared = await mutate(entries => { entries[rels] = strToU8(strFromU8(entries[rels]).replace('Target="../media/track2.vtt"', 'Target="../media/track1.vtt"')); });
  const result = await inspectPptxCaptionTracks(shared, { captionLimits: { maxTotalTrackBytes: expected.tracks[0].byteLength + expected.tracks[2].byteLength } });
  assert.equal(result.inspectedBytes, expected.tracks[0].byteLength + expected.tracks[2].byteLength); assert.equal(result.tracks[0].targetPart, result.tracks[1].targetPart);
  const ac = new AbortController(); ac.abort(); await assert.rejects(inspectPptxCaptionTracks(bytes, { signal: ac.signal }), error => error.code === 'operation-aborted');
});

test('raw source inventory includes but does not activate an inactive MCE choice', async () => {
  const bytes = await mutate(entries => {
    const key = 'ppt/slides/slide1.xml'; let xml = strFromU8(entries[key]);
    xml = xml.replace('<p:pic>', '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:f="urn:unsupported"><mc:Choice Requires="f"><p:pic>');
    xml = xml.replace('</p:pic>', '</p:pic></mc:Choice><mc:Fallback/></mc:AlternateContent>'); entries[key] = strToU8(xml);
  });
  const result = await inspectPptxCaptionTracks(bytes);
  assert.deepEqual(result.tracks.map(x => x.objectId), ['2', '2', '3']); assert.ok(result.tracks[0].source.path.includes('AlternateContent'));
  const session = await importPptx(bytes, { projection: 'content' });
  try { const plan = await session.prepareSave({ mode: 'no-op' }); assert.deepEqual((await session.save(plan.plan)).bytes, bytes); } finally { session.dispose(); }
});
