import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile, writeFile, mkdtemp, rm } from 'node:fs/promises';
import { execFileSync } from 'node:child_process';
import { createRequire } from 'node:module';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { importPptx } from '../dist/index.js';

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
