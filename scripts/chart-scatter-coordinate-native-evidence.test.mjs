import assert from 'node:assert/strict';
import test from 'node:test';
import { execFileSync } from 'node:child_process';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { importPptx } from '../packages/pptx-parser/dist/index.js';

const base = '.tmp/native-scatter-coordinate-save/powerpoint-bI3pt5/';
const record = 'docs/testing/chart-scatter-coordinate-save-powerpoint-evidence.json';
const verify = file => execFileSync('python3', ['scripts/research/verify-scatter-coordinate-save-powerpoint.py', ...(file ? [file] : [])], { encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'] });

test('PowerPoint opens, saves and reopens the source-bound scatter coordinate edits', () => {
  assert.deepEqual(JSON.parse(verify()), { verified: true, charts: 1, coordinateEdits: 3, reopenRounds: 1 });
});

test('native-saved dense coordinates remain independently bound for a later edit', async () => {
  const bytes = await readFile(base + 'native-round-1.pptx');
  const session = await importPptx(bytes, { projection: 'content', chartEditSaveProfile: 'existing-values-v1' });
  try {
    assert.equal(session.report.chartValues.length, 12);
    assert.ok(session.report.chartValues.every(binding => binding.status === 'bound'));
    const slide = session.document.slides[0], chart = structuredClone(slide.elements.find(element => element.type === 'chart'));
    assert.deepEqual(chart.series.map(series => [series.xValues, series.yValues]), [[[-3, 2, 8], [8, -4, 6]], [[7, -1, 3], [-6, -7, 2]]]);
    chart.series[0].xValues[0] = -2; chart.series[1].yValues[1] = -6;
    slide.replaceElement(chart.id, chart);
    const plan = await session.prepareSave({ mode: 'edit' }); assert.equal(plan.status, 'ready', JSON.stringify(plan.report));
    const reopened = await importPptx((await session.save(plan.plan)).bytes, { projection: 'content' });
    try { const result = reopened.document.slides[0].elements.find(element => element.type === 'chart'); assert.equal(result.series[0].xValues[0], -2); assert.equal(result.series[1].yValues[1], -6); }
    finally { reopened.dispose(); }
  } finally { session.dispose(); }
});

test('scatter coordinate evidence rejects altered package and capture identities', async () => {
  const original = JSON.parse(await readFile(record, 'utf8')), dir = await mkdtemp(path.join(tmpdir(), 'pptkit-scatter-coordinate-evidence-'));
  try {
    for (const mutate of [r => { r.packages[2].sha256 = '0'.repeat(64); }, r => { r.captures[1].axSha256 = '0'.repeat(64); }, r => { r.captures[0].screenshotSha256 = '0'.repeat(64); }, r => { r.scope = 'changed'; }]) {
      const bad = structuredClone(original); mutate(bad); const file = path.join(dir, 'bad.json'); await writeFile(file, JSON.stringify(bad)); assert.throws(() => verify(file));
    }
  } finally { await rm(dir, { recursive: true, force: true }); }
});

test('disabled native Edit Data observation covers both samples and rejects altered identities', async () => {
  const command = 'scripts/research/verify-scatter-edit-data-availability.py', evidence = 'docs/testing/chart-scatter-edit-data-availability.json';
  const run = file => execFileSync('python3', [command, ...(file ? [file] : [])], { encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'] });
  assert.deepEqual(JSON.parse(run()), { verified: true, disabledControls: 2, interactiveEditPerformed: false });
  const original = JSON.parse(await readFile(evidence, 'utf8')), dir = await mkdtemp(path.join(tmpdir(), 'pptkit-scatter-edit-availability-'));
  try { for (const mutate of [r => { r.captures[0].fileSha256 = '0'.repeat(64); }, r => { r.captures[1].axSha256 = '0'.repeat(64); }, r => { r.captures[0].screenshotSha256 = '0'.repeat(64); }, r => { r.scope = 'changed'; }]) { const bad = structuredClone(original); mutate(bad); const file = path.join(dir, 'bad.json'); await writeFile(file, JSON.stringify(bad)); assert.throws(() => run(file)); } }
  finally { await rm(dir, { recursive: true, force: true }); }
});
