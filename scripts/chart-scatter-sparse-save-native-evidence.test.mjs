import assert from 'node:assert/strict';
import test from 'node:test';
import { execFileSync } from 'node:child_process';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { importPptx } from '../packages/pptx-parser/dist/index.js';

const base = '.tmp/native-scatter-sparse-save/powerpoint-CbeAdy/', evidence = 'docs/testing/chart-scatter-sparse-save-powerpoint-evidence.json';
const verify = file => execFileSync('python3', ['scripts/research/verify-scatter-sparse-save-powerpoint.py', ...(file ? [file] : [])], { encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'] });

test('PowerPoint opens and reopens four sparse coordinate save samples', () => {
  assert.deepEqual(JSON.parse(verify()), { verified: true, charts: 4, coordinateEdits: 8, reopenRounds: 1 });
});

test('PowerPoint-saved sparse sources remain bound and support a subsequent source edit', async () => {
  const bytes = await readFile(base + 'native-round-1.pptx'), session = await importPptx(bytes, { projection: 'content', chartEditSaveProfile: 'existing-values-v2' });
  try {
    assert.equal(session.report.chartValues.length, 60); assert.ok(session.report.chartValues.every(binding => binding.status === 'bound'));
    for (const slide of session.document.slides) { const chart = structuredClone(slide.elements.find(e => e.type === 'chart')); assert.deepEqual(chart.series[0].yValues, [4, null, -3, 6]); chart.series[0].xValues[0] = -5; slide.replaceElement(chart.id, chart); }
    const plan = await session.prepareSave({ mode: 'edit' }); assert.equal(plan.status, 'ready', JSON.stringify(plan.report));
    const reopened = await importPptx((await session.save(plan.plan)).bytes, { projection: 'content' });
    try { for (const slide of reopened.document.slides) { const chart = slide.elements.find(e => e.type === 'chart'); assert.equal(chart.series[0].xValues[0], -5); assert.equal(chart.series[0].yValues[1], null); } } finally { reopened.dispose(); }
  } finally { session.dispose(); }
});

test('sparse save evidence rejects changed package and UI capture hashes', async () => {
  const original = JSON.parse(await readFile(evidence, 'utf8')), dir = await mkdtemp(path.join(tmpdir(), 'pptkit-sparse-native-'));
  try { for (const mutate of [r => { r.packages[2].sha256 = '0'.repeat(64); }, r => { r.captures[0].axSha256 = '0'.repeat(64); }, r => { r.captures[7].screenshotSha256 = '0'.repeat(64); }, r => { r.scope = 'changed'; }]) { const bad = structuredClone(original); mutate(bad); const file = path.join(dir, 'bad.json'); await writeFile(file, JSON.stringify(bad)); assert.throws(() => verify(file)); } }
  finally { await rm(dir, { recursive: true, force: true }); }
});
