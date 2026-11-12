import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { createPresentation } from '@pptkit/core';
import { generatePptx } from '@pptkit/pptx-exporter';
import { importPptx } from '@pptkit/pptx-parser';
import { runCliAsync } from '../dist/index.js';

async function fixture(t) {
  const dir = await mkdtemp(path.join(tmpdir(), 'pptkit-cli-scatter-')); t.after(() => rm(dir, { recursive: true, force: true }));
  const p = createPresentation(), slide = p.addSlide();
  slide.addElement({ type: 'chart', chartType: 'scatter', box: { x: 0, y: 0, width: 300, height: 200 }, series: [{ name: 'S', xValues: [1, 2], yValues: [3, 4] }] });
  slide.addElement({ type: 'chart', chartType: 'bar', box: { x: 300, y: 0, width: 300, height: 200 }, categories: ['A'], series: [{ name: 'B', values: [5] }] });
  await writeFile(path.join(dir, 'source.pptx'), (await generatePptx(p)).bytes);
  return dir;
}
const call = (dir, args) => runCliAsync([...args, '--json'], { cwd: dir });
async function inspect(dir, file, profile = 'existing-values-v1') {
  const result = await call(dir, ['inspect', file, '--projection', 'content', '--chart-edit-save-profile', profile, '--chart-values']);
  assert.equal(result.exitCode, 0, result.output); return JSON.parse(result.output);
}

test('v2 CLI edits X, Y and category values atomically over three reopens', async t => {
  const dir = await fixture(t); let input = 'source.pptx';
  for (let round = 0; round < 3; round++) {
    const report = await inspect(dir, input), v1 = report.chartEditCandidates, v2 = report.chartEditCandidatesV2;
    assert.equal(v1.version, 'pptkit-chart-edits-v1'); assert.equal(v1.targets.length, 1); assert.equal(v1.unboundCount, 4);
    assert.equal(v2.version, 'pptkit-chart-edits-v2'); assert.equal(v2.targets.length, 5); assert.equal(v2.unboundCount, 0);
    const x = v2.targets.find(v => v.coordinate === 'x' && v.pointIndex === 0), y = v2.targets.find(v => v.coordinate === 'y' && v.pointIndex === 0), bar = v2.targets.find(v => v.coordinate === undefined);
    const manifest = { version: v2.version, sourceSha256: v2.sourceSha256, edits: [{ ...x, replacement: String(10 + round) }, { ...y, replacement: String(20 + round) }, { ...bar, replacement: String(30 + round) }] };
    await writeFile(path.join(dir, 'edit.json'), JSON.stringify(manifest)); const output = `round-${round}.pptx`;
    const result = await call(dir, ['save', input, output, '--mode', 'edit', '--chart-edits', 'edit.json']); assert.equal(result.exitCode, 0, result.output);
    const s = await importPptx(await readFile(path.join(dir, output)), { projection: 'content' });
    try { const [scatter, category] = s.document.slides[0].elements; assert.equal(scatter.series[0].xValues[0], 10 + round); assert.equal(scatter.series[0].yValues[0], 20 + round); assert.equal(category.series[0].values[0], 30 + round); } finally { s.dispose(); }
    input = output;
  }
});

test('v3 CLI preserves sparse blanks and edits existing X/Y and category values over three rounds', async t => {
  const dir = await fixture(t), p = createPresentation(), slide = p.addSlide();
  slide.addElement({ type: 'chart', chartType: 'scatter', box: { x: 0, y: 0, width: 300, height: 200 }, blankDisplay: 'span', series: [{ name: 'S', xValues: [1, null, 3], yValues: [4, 5, 6] }] });
  slide.addElement({ type: 'chart', chartType: 'bar', box: { x: 300, y: 0, width: 300, height: 200 }, categories: ['A'], series: [{ name: 'B', values: [5] }] });
  await writeFile(path.join(dir, 'sparse.pptx'), (await generatePptx(p)).bytes); let input = 'sparse.pptx';
  for (let round = 0; round < 3; round++) {
    const report = await inspect(dir, input, 'existing-values-v2'), candidates = report.chartEditCandidatesV3;
    assert.equal(report.chartEditCandidates, undefined); assert.equal(report.chartEditCandidatesV2, undefined); assert.equal(candidates.version, 'pptkit-chart-edits-v3'); assert.equal(candidates.targets.length, 6); assert.equal(candidates.unboundCount, 0);
    assert.equal(candidates.targets.some(target => target.coordinate === 'x' && target.pointIndex === 1), false);
    const x = candidates.targets.find(target => target.coordinate === 'x' && target.pointIndex === 2), y = candidates.targets.find(target => target.coordinate === 'y' && target.pointIndex === 0), bar = candidates.targets.find(target => target.coordinate === undefined);
    await writeFile(path.join(dir, 'v3.json'), JSON.stringify({ version: candidates.version, sourceSha256: candidates.sourceSha256, edits: [{ ...x, replacement: String(7 + round) }, { ...y, replacement: String(10 + round) }, { ...bar, replacement: String(20 + round) }] }));
    const output = `sparse-${round}.pptx`, result = await call(dir, ['save', input, output, '--mode', 'edit', '--chart-edits', 'v3.json']); assert.equal(result.exitCode, 0, result.output);
    const reopened = await importPptx(await readFile(path.join(dir, output)), { projection: 'content', chartEditSaveProfile: 'existing-values-v2' });
    try { const [scatter, category] = reopened.document.slides[0].elements; assert.deepEqual(scatter.series[0].xValues, [1, null, 7 + round]); assert.equal(scatter.series[0].yValues[0], 10 + round); assert.equal(category.series[0].values[0], 20 + round); } finally { reopened.dispose(); }
    input = output;
  }
});

test('v3 rejects blank, stale and mixed invalid edits without publishing; v2 remains restrictive', async t => {
  const dir = await fixture(t), p = createPresentation(); p.addSlide().addElement({ type: 'chart', chartType: 'scatter', box: { x: 0, y: 0, width: 300, height: 200 }, blankDisplay: 'gap', series: [{ name: 'S', xValues: [1, null, 3], yValues: [4, 5, 6] }] });
  const bytes = (await generatePptx(p)).bytes; await writeFile(path.join(dir, 'sparse.pptx'), bytes);
  const candidates = (await inspect(dir, 'sparse.pptx', 'existing-values-v2')).chartEditCandidatesV3, x = candidates.targets.find(target => target.coordinate === 'x' && target.pointIndex === 0), edit = { ...x, replacement: '9' };
  const cases = [ { edits: [{ ...edit, pointIndex: 1 }] }, { edits: [edit, edit] }, { edits: [edit, { ...edit, coordinate: 'y', expected: '999' }] }, { edits: [{ ...edit, replacement: 'NaN' }] }, { edits: [edit], sourceSha256: '0'.repeat(64) }, { edits: [edit], version: 'pptkit-chart-edits-v2' } ];
  for (const [index, bad] of cases.entries()) { const manifest = { version: candidates.version, sourceSha256: candidates.sourceSha256, ...bad }; await writeFile(path.join(dir, 'bad-v3.json'), JSON.stringify(manifest)); const output = `bad-v3-${index}.pptx`, result = await call(dir, ['save', 'sparse.pptx', output, '--mode', 'edit', '--chart-edits', 'bad-v3.json']); assert.equal(result.exitCode, 3, result.output); await assert.rejects(readFile(path.join(dir, output)), { code: 'ENOENT' }); assert.deepEqual(await readFile(path.join(dir, 'sparse.pptx')), Buffer.from(bytes)); }
});

test('v3 sparse chart edits combine atomically with text and slide numbering', async t => {
  const dir = await fixture(t), p = createPresentation(), slide = p.addSlide();
  slide.addElement({ type: 'chart', chartType: 'scatter', box: { x: 0, y: 0, width: 300, height: 200 }, blankDisplay: 'gap', series: [{ name: 'S', xValues: [1, null, 3], yValues: [4, 5, 6] }] });
  slide.addElement({ type: 'text', box: { x: 0, y: 210, width: 300, height: 60 }, content: 'before' });
  await writeFile(path.join(dir, 'mixed-source.pptx'), (await generatePptx(p)).bytes);
  const inspection = await call(dir, ['inspect', 'mixed-source.pptx', '--projection', 'content', '--edit-save-profile', 'text-runs-v3', '--text-values', '--chart-edit-save-profile', 'existing-values-v2', '--chart-values']); assert.equal(inspection.exitCode, 0, inspection.output);
  const report = JSON.parse(inspection.output), chart = report.chartEditCandidatesV3, text = report.textEditCandidates;
  const coordinate = chart.targets.find(target => target.coordinate === 'x' && target.pointIndex === 2), run = text.targets.find(target => target.expected === 'before'); assert.ok(coordinate && run);
  await writeFile(path.join(dir, 'mixed-chart.json'), JSON.stringify({ version: chart.version, sourceSha256: chart.sourceSha256, edits: [{ ...coordinate, replacement: '8' }] }));
  await writeFile(path.join(dir, 'mixed-text.json'), JSON.stringify({ version: text.version, sourceSha256: text.sourceSha256, edits: [{ ...run, replacement: 'after' }] }));
  const output = await call(dir, ['save', 'mixed-source.pptx', 'mixed-output.pptx', '--mode', 'edit', '--chart-edits', 'mixed-chart.json', '--text-edits', 'mixed-text.json', '--first-slide-number', '17']); assert.equal(output.exitCode, 0, output.output);
  const reopened = await importPptx(await readFile(path.join(dir, 'mixed-output.pptx')), { projection: 'content' }); try { assert.equal(reopened.document.firstSlideNumber, 17); assert.deepEqual(reopened.document.slides[0].elements[0].series[0].xValues, [1, null, 8]); assert.equal(reopened.document.slides[0].elements[1].content[0].runs[0].text, 'after'); } finally { reopened.dispose(); }
});

test('v2 rejects ambiguous, malformed, stale and unbound targets without publishing', async t => {
  const dir = await fixture(t), original = await readFile(path.join(dir, 'source.pptx')), report = await inspect(dir, 'source.pptx'), candidates = report.chartEditCandidatesV2;
  const x = candidates.targets.find(v => v.coordinate === 'x' && v.pointIndex === 0), bar = candidates.targets.find(v => v.coordinate === undefined), edit = { ...x, replacement: '9' };
  const cases = [
    [{ ...edit, coordinate: 'z' }], [{ ...edit, coordinate: null }], [{ ...edit, coordinate: undefined }],
    [{ ...edit, expected: '99' }], [edit, edit], [{ ...bar, coordinate: 'x', replacement: '9' }],
    [{ ...edit, replacement: '01' }], [{ ...edit, replacement: '9' }, { ...bar, expected: '99', replacement: '7' }],
  ];
  for (const [index, edits] of cases.entries()) {
    await writeFile(path.join(dir, 'bad.json'), JSON.stringify({ version: candidates.version, sourceSha256: candidates.sourceSha256, edits }));
    const output = `bad-${index}.pptx`, result = await call(dir, ['save', 'source.pptx', output, '--mode', 'edit', '--chart-edits', 'bad.json']);
    assert.equal(result.exitCode, 3, result.output); await assert.rejects(readFile(path.join(dir, output)), { code: 'ENOENT' }); assert.deepEqual(await readFile(path.join(dir, 'source.pptx')), original);
  }
  await writeFile(path.join(dir, 'bad.json'), JSON.stringify({ version: 'pptkit-chart-edits-v1', sourceSha256: candidates.sourceSha256, edits: [edit] }));
  assert.equal((await call(dir, ['save', 'source.pptx', 'v1-invalid.pptx', '--mode', 'edit', '--chart-edits', 'bad.json'])).exitCode, 3);
  await writeFile(path.join(dir, 'bad.json'), JSON.stringify({ version: candidates.version, sourceSha256: '0'.repeat(64), edits: [edit] }));
  const stale = await call(dir, ['save', 'source.pptx', 'stale.pptx', '--mode', 'edit', '--chart-edits', 'bad.json']); assert.equal(stale.exitCode, 3); assert.equal(JSON.parse(stale.output).error.code, 'chart-edits-source-mismatch');
  await assert.rejects(readFile(path.join(dir, 'stale.pptx')), { code: 'ENOENT' });
});

test('v2 reports sparse scatter coordinates as unbound', async t => {
  const dir = await fixture(t), p = createPresentation(), slide = p.addSlide();
  slide.addElement({ type: 'chart', chartType: 'scatter', box: { x: 0, y: 0, width: 300, height: 200 }, blankDisplay: 'gap', series: [{ name: 'S', xValues: [1, null], yValues: [3, 4] }] });
  await writeFile(path.join(dir, 'sparse.pptx'), (await generatePptx(p)).bytes);
  const report = await inspect(dir, 'sparse.pptx'); assert.equal(report.chartEditCandidatesV2.targets.length, 0); assert.equal(report.chartEditCandidatesV2.unboundCount, 3);
});
