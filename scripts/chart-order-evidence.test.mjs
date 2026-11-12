import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { spawnSync } from 'node:child_process';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { importPptx } from '../packages/pptx-parser/dist/index.js';
import { readGeneratedZip as readGeneratedZipForEvidence } from '../tests/support/read-generated-zip.mjs';
import { replaceArchiveEntries } from '../packages/pptx-parser/dist/preservation/replace-archive-entries.js';

test('independent ZIP and XML oracle checks retained PowerPoint counterexample', async () => {
  const before = 'tests/fixtures/chart-label-source.pptx', after = 'tests/fixtures/chart-order-powerpoint-round-1.pptx';
  assert.equal(createHash('sha256').update(await readFile(after)).digest('hex'), 'ab78ff44368fea9f05a4a11c38bedeec8a94d0493f28b657a3813f9c19b2b20c');
  const unchanged = spawnSync('python3', ['scripts/research/verify-chart-series-order.py', before, before], { encoding: 'utf8' });
  assert.notEqual(unchanged.status, 0, 'oracle must reject an unchanged order');
  const result = spawnSync('python3', ['scripts/research/verify-chart-series-order.py', before, after], { encoding: 'utf8' });
  assert.equal(result.status, 0, result.stderr);
});

test('physical series swap controls preserve payloads and expose Parser ordering difference', async () => {
  const fixtures = [
    ['tests/fixtures/chart-label-source.pptx', 'tests/fixtures/chart-order-powerpoint-move-only.pptx', 'acf0189959a55a9ee4531fbaabc26a1a81b51b59a1c2b3a02ac9e781174c588b', ['First _x005F_', 'Second & series']],
    ['tests/fixtures/chart-order-powerpoint-round-1.pptx', 'tests/fixtures/chart-order-powerpoint-move-and-order.pptx', 'bf27735149a585f07414e8b98053833705291935d68296b701cff29a43f75797', ['Second & series', 'First _x005F_']],
  ];
  for (const [before, after, hash, names] of fixtures) {
    const bytes = await readFile(after); assert.equal(createHash('sha256').update(bytes).digest('hex'), hash);
    const result = spawnSync('python3', ['scripts/research/verify-chart-series-move.py', before, after], { encoding: 'utf8' }); assert.equal(result.status, 0, result.stderr);
    const unchanged = spawnSync('python3', ['scripts/research/verify-chart-series-move.py', before, before], { encoding: 'utf8' }); assert.notEqual(unchanged.status, 0);
    const session = await importPptx(bytes, { projection: 'content' });
    try { assert.deepEqual(session.document.slides[0].elements[0].series.map(series => series.name), names); assert.equal(session.report.findings.some(item => item.feature === 'chart.series-order'), before.endsWith('chart-label-source.pptx')); } finally { session.dispose(); }
  }
});

test('PowerPoint-opened mixed order candidates pass independent three-round ZIP/XML checks', async () => {
  const hashes = ['05ab484c0bc123bcb1dfc334e95db8d89758ecda0d36c580552d6d6ad3bcf1b5', 'fb93b4394687eb2a1e98077d9889b077e338c0bed71a58f50c1da1c83b58de56', 'e2dfc8287adb148fbed1414e63e8f7b2ec03ba828a32d1f1a15ab8a008e76886'];
  let before = 'tests/fixtures/chart-label-source.pptx';
  for (const [index, hash] of hashes.entries()) {
    const after = `tests/fixtures/chart-order-candidate-round-${index + 1}.pptx`;
    assert.equal(createHash('sha256').update(await readFile(after)).digest('hex'), hash);
    const args = ['scripts/research/verify-chart-order-candidate.py', before, after, String(index + 1)];
    const result = spawnSync('python3', args, { encoding: 'utf8' }); assert.equal(result.status, 0, result.stderr);
    const wrong = spawnSync('python3', [args[0], before, before, args[3]], { encoding: 'utf8' }); assert.notEqual(wrong.status, 0);
    before = after;
  }
});

test('PowerPoint legend dependency controls contain only the indexed legend override', async () => {
  const entries = [
    ['tests/fixtures/chart-label-source.pptx', 'tests/fixtures/chart-order-legend-before.pptx', '1bd0d951cd265de8dbc48f47047bd9d2056f6140bd6a6b52e2ff9778bf6a65f8'],
    ['tests/fixtures/chart-order-powerpoint-move-and-order.pptx', 'tests/fixtures/chart-order-legend-after.pptx', 'ae593263c1814b0ad23f3f88d5763e033c014a6f367a6d401104cf97ee875a67'],
  ];
  for (const [base, file, hash] of entries) {
    assert.equal(createHash('sha256').update(await readFile(file)).digest('hex'), hash);
    const result = spawnSync('python3', ['scripts/research/verify-chart-legend-control.py', base, file], { encoding: 'utf8' }); assert.equal(result.status, 0, result.stderr);
    const negative = spawnSync('python3', ['scripts/research/verify-chart-legend-control.py', base, base], { encoding: 'utf8' }); assert.notEqual(negative.status, 0);
  }
});

test('remapped PowerPoint legend control changes only the selected legend index', async () => {
  const file = 'tests/fixtures/chart-order-legend-remapped.pptx', bytes = await readFile(file);
  assert.equal(createHash('sha256').update(bytes).digest('hex'), 'f045dc1eda30ae008b431fb832b7a88b98675904427e28d4a9ea483f50bf3c7d');
  const before = readGeneratedZipForEvidence(await readFile('tests/fixtures/chart-order-legend-after.pptx')), after = readGeneratedZipForEvidence(bytes);
  assert.deepEqual([...before.keys()].sort(), [...after.keys()].sort());
  for (const [name, value] of before) {
    if (name === 'ppt/charts/chart1.xml') {
      const old = value.toString(), changed = old.replace('<c:legendEntry><c:idx val="0"/>', '<c:legendEntry><c:idx val="1"/>');
      assert.notEqual(changed, old); assert.equal(after.get(name).toString(), changed);
    } else assert.deepEqual(after.get(name), value, name);
  }
});

test('three PowerPoint legend candidates pass an independent mixed-edit oracle', async t => {
  const hashes = ['44c4c769ed152628c3fdd4f7b0e5c132f82cced313791b573f57bc4186b3a009', '0b7cba231cbb635878a015f6f2a64ac6bcc82140ef7b2ed9ec5da5a1637ccaa5', '00df05a63a41cc73d44f9a699388e2d60802ec151b716c07ab12f05100dad876'];
  let before = 'tests/fixtures/chart-order-legend-before.pptx';
  for (const [index, hash] of hashes.entries()) {
    const file = `tests/fixtures/chart-order-legend-candidate-round-${index + 1}.pptx`;
    assert.equal(createHash('sha256').update(await readFile(file)).digest('hex'), hash);
    const args = ['scripts/research/verify-chart-order-candidate.py', before, file, String(index + 1), 'legend'];
    const result = spawnSync('python3', args, { encoding: 'utf8' }); assert.equal(result.status, 0, result.stderr);
    const wrong = spawnSync('python3', [args[0], before, before, args[3], 'legend'], { encoding: 'utf8' }); assert.notEqual(wrong.status, 0);
    before = file;
  }
  const directory = await mkdtemp(path.join(tmpdir(), 'pptkit-legend-oracle-')); t.after(() => rm(directory, { recursive: true, force: true }));
  const bytes = await readFile('tests/fixtures/chart-order-legend-candidate-round-1.pptx'), name = 'ppt/charts/chart1.xml', xml = readGeneratedZipForEvidence(bytes).get(name).toString();
  const corrupt = xml.replace('<c:legendEntry><c:idx val="1"/>', '<c:legendEntry><c:idx val="0"/>'); assert.notEqual(corrupt, xml);
  const wrong = await replaceArchiveEntries(bytes, [{ name, bytes: new TextEncoder().encode(corrupt) }], { maxOutputBytes: 10_000_000 });
  const wrongPath = path.join(directory, 'wrong.pptx'); await writeFile(wrongPath, wrong);
  const result = spawnSync('python3', ['scripts/research/verify-chart-order-candidate.py', 'tests/fixtures/chart-order-legend-before.pptx', wrongPath, '1', 'legend'], { encoding: 'utf8' }); assert.notEqual(result.status, 0);
});

test('PowerPoint resave retains chart, workbook and hidden legend correspondence', async () => {
  const file = 'tests/fixtures/chart-order-legend-powerpoint-resave.pptx', bytes = await readFile(file);
  assert.equal(createHash('sha256').update(bytes).digest('hex'), 'e78673533a819b8f8ed384928d79cce5b5d3a3d427967469d4a19ac32c6f0d05');
  const result = spawnSync('python3', ['scripts/research/verify-powerpoint-order-legend-resave.py', file], { encoding: 'utf8' }); assert.equal(result.status, 0, result.stderr);
  const negative = spawnSync('python3', ['scripts/research/verify-powerpoint-order-legend-resave.py', 'tests/fixtures/chart-order-legend-before.pptx'], { encoding: 'utf8' }); assert.notEqual(negative.status, 0);
});

test('PowerPoint series-extension move control retains extension bytes and ownership', async () => {
  const before = 'tests/fixtures/chart-order-legend-powerpoint-resave.pptx', after = 'tests/fixtures/chart-order-powerpoint-extensions-moved.pptx';
  assert.equal(createHash('sha256').update(await readFile(after)).digest('hex'), 'e006eb02b16786ee30d27066f338b74831103aa355438b54ce65b8c37d78c607');
  const result = spawnSync('python3', ['scripts/research/verify-chart-extension-move.py', before, after], { encoding: 'utf8' }); assert.equal(result.status, 0, result.stderr);
  const wrong = spawnSync('python3', ['scripts/research/verify-chart-extension-move.py', before, before], { encoding: 'utf8' }); assert.notEqual(wrong.status, 0);
});
