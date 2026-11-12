import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { importPptx } from '@pptkit/pptx-parser';
import { runCliAsync } from '../dist/index.js';

test('v5 CLI discovers and atomically saves shared labels and numbers across three reopens', async t => {
  const dir = await mkdtemp(path.join(tmpdir(), 'pptkit-label-cli-')); t.after(() => rm(dir, { recursive: true, force: true }));
  await writeFile(path.join(dir, 'source.pptx'), await readFile(new URL('../../../tests/fixtures/chart-label-source.pptx', import.meta.url)));
  const call = args => runCliAsync([...args, '--json'], { cwd: dir }); let input = 'source.pptx';
  for (const [round, value] of ['changed', ' & _x0041_ ', '中文 😀'].entries()) {
    const inspected = await call(['inspect', input, '--projection', 'content', '--chart-edit-save-profile', 'existing-values-and-labels-v4', '--chart-values']);
    assert.equal(inspected.exitCode, 0, inspected.output); const result = JSON.parse(inspected.output), candidates = result.chartEditCandidatesV5;
    assert.equal(candidates.version, 'pptkit-chart-edits-v5'); assert.equal(candidates.unboundCount, 0);
    const category = candidates.targets.find(item => item.kind === 'category' && item.categoryIndex === 1);
    const numeric = candidates.targets.find(item => item.kind === undefined && item.seriesIndex === 0 && item.pointIndex === 0);
    assert.ok(category && numeric); const edit = { ...category, replacement: value };
    const manifest = { version: candidates.version, sourceSha256: candidates.sourceSha256, edits: [edit, { ...numeric, replacement: String(40 + round) }] };
    await writeFile(path.join(dir, 'edits.json'), JSON.stringify(manifest)); const output = `round-${round}.pptx`;
    const saved = await call(['save', input, output, '--mode', 'edit', '--chart-edits', 'edits.json']); assert.equal(saved.exitCode, 0, saved.output);
    const session = await importPptx(await readFile(path.join(dir, output)), { projection: 'content' });
    try { const chart = session.document.slides[0].elements[0]; assert.equal(chart.categories[1], value); assert.equal(chart.series[0].values[0], 40 + round); } finally { session.dispose(); }
    for (const [index, invalid] of [[edit, edit], [{ ...edit, expected: 'stale' }], [{ ...edit, replacement: '' }], [{ ...edit, categoryIndex: 99 }]].entries()) {
      await writeFile(path.join(dir, 'bad.json'), JSON.stringify({ ...manifest, edits: invalid })); const rejected = await call(['save', input, `bad-${round}-${index}.pptx`, '--mode', 'edit', '--chart-edits', 'bad.json']);
      assert.equal(rejected.exitCode, 3, rejected.output); await assert.rejects(readFile(path.join(dir, `bad-${round}-${index}.pptx`)), { code: 'ENOENT' });
    }
    input = output;
  }
});

test('v6 CLI saves series names, category labels and numbers together with guarded manifests', async t => {
  const dir = await mkdtemp(path.join(tmpdir(), 'pptkit-series-cli-')); t.after(() => rm(dir, { recursive: true, force: true }));
  await writeFile(path.join(dir, 'source.pptx'), await readFile(new URL('../../../tests/fixtures/chart-label-source.pptx', import.meta.url)));
  const call = args => runCliAsync([...args, '--json'], { cwd: dir }); let input = 'source.pptx';
  for (const [round, value] of ['name & text', ' _x0041_ ', '中文 😀'].entries()) {
    const inspected = await call(['inspect', input, '--projection', 'content', '--chart-edit-save-profile', 'existing-values-and-series-names-v5', '--chart-values']);
    assert.equal(inspected.exitCode, 0, inspected.output); const candidates = JSON.parse(inspected.output).chartEditCandidatesV6;
    assert.equal(candidates.version, 'pptkit-chart-edits-v6'); assert.equal(candidates.unboundCount, 0);
    const name = candidates.targets.find(item => item.kind === 'series-name' && item.seriesIndex === 0), category = candidates.targets.find(item => item.kind === 'category' && item.categoryIndex === 1), numeric = candidates.targets.find(item => !item.kind && item.seriesIndex === 0 && item.pointIndex === 0);
    assert.ok(name && category && numeric);
    const edit = { ...name, replacement: value }, manifest = { version: candidates.version, sourceSha256: candidates.sourceSha256, edits: [edit, { ...category, replacement: `category ${round}` }, { ...numeric, replacement: String(90 + round) }] };
    await writeFile(path.join(dir, 'edit.json'), JSON.stringify(manifest)); const output = `round-${round}.pptx`;
    const saved = await call(['save', input, output, '--mode', 'edit', '--chart-edits', 'edit.json']); assert.equal(saved.exitCode, 0, saved.output);
    const session = await importPptx(await readFile(path.join(dir, output)), { projection: 'content' });
    try { const chart = session.document.slides[0].elements[0]; assert.equal(chart.series[0].name, value); assert.equal(chart.categories[1], `category ${round}`); assert.equal(chart.series[0].values[0], 90 + round); } finally { session.dispose(); }
    const invalids = [[edit, edit], [{ ...edit, expected: 'stale' }], [{ ...edit, replacement: '' }], [{ ...edit, seriesIndex: 99 }]];
    for (const [index, edits] of invalids.entries()) { await writeFile(path.join(dir, 'bad.json'), JSON.stringify({ ...manifest, edits })); const result = await call(['save', input, `bad-${round}-${index}.pptx`, '--mode', 'edit', '--chart-edits', 'bad.json']); assert.equal(result.exitCode, 3, result.output); await assert.rejects(readFile(path.join(dir, `bad-${round}-${index}.pptx`)), { code: 'ENOENT' }); }
    await writeFile(path.join(dir, 'bad.json'), JSON.stringify({ ...manifest, version: 'pptkit-chart-edits-v5', edits: [edit] }));
    assert.equal((await call(['save', input, `old-${round}.pptx`, '--mode', 'edit', '--chart-edits', 'bad.json'])).exitCode, 3);
    input = output;
  }
});

test('v6 CLI combines series names with scatter and bubble coordinates', async t => {
  const dir = await mkdtemp(path.join(tmpdir(), 'pptkit-coordinate-name-cli-')); t.after(() => rm(dir, { recursive: true, force: true }));
  const call = args => runCliAsync([...args, '--json'], { cwd: dir });
  for (const family of ['scatter', 'bubble']) {
    await writeFile(path.join(dir, `${family}.pptx`), await readFile(new URL(`../../../tests/fixtures/${family}-name-source.pptx`, import.meta.url)));
    const inspection = await call(['inspect', `${family}.pptx`, '--projection', 'content', '--chart-edit-save-profile', 'existing-values-and-series-names-v5', '--chart-values']); assert.equal(inspection.exitCode, 0, inspection.output);
    const candidates = JSON.parse(inspection.output).chartEditCandidatesV6, name = candidates.targets.find(item => item.kind === 'series-name'), point = candidates.targets.find(item => item.coordinate === (family === 'bubble' ? 'size' : 'x') && item.pointIndex === 0);
    assert.ok(name && point); const manifest = { version: candidates.version, sourceSha256: candidates.sourceSha256, edits: [{ ...name, replacement: 'renamed' }, { ...point, replacement: '8' }] };
    await writeFile(path.join(dir, 'edit.json'), JSON.stringify(manifest)); const result = await call(['save', `${family}.pptx`, `${family}-out.pptx`, '--mode', 'edit', '--chart-edits', 'edit.json']); assert.equal(result.exitCode, 0, result.output);
    const session = await importPptx(await readFile(path.join(dir, `${family}-out.pptx`)), { projection: 'content' });
    try { const series = session.document.slides[0].elements[0].series[0]; assert.equal(series.name, 'renamed'); assert.equal(family === 'bubble' ? series.sizes[0] : series.xValues[0], 8); } finally { session.dispose(); }
  }
});
