import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { execFileSync } from 'node:child_process';
import path from 'node:path';
import { tmpdir } from 'node:os';
import { readGeneratedZip } from '../tests/support/read-generated-zip.mjs';
import { zip } from '../packages/pptx-parser/test/zip-fixture.mjs';

const python = (...args) => execFileSync('python3', ['scripts/research/verify-combined-number-chart.py', ...args], { encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'] });
const makeZip = parts => zip([...parts].map(([name, body]) => ({ name, body }))).bytes;

test('independent combined-save oracle rejects wrong title, number, chart, workbook and unrelated changes', async t => {
  const parent = await mkdtemp(path.join(tmpdir(), 'pptkit-combined-evidence-')); t.after(() => rm(parent, { recursive: true, force: true }));
  const directory = path.join(parent, 'sample'); execFileSync(process.execPath, ['scripts/prepare-combined-number-chart-native.mjs', directory]);
  const source = path.join(directory, 'source.pptx'), candidate = path.join(directory, 'candidate.pptx');
  const evidence = JSON.parse(python(source, candidate)); assert.equal(evidence.first_slide_number, 17); assert.equal(evidence.chart_value, 7); assert.equal(evidence.cell, 'B3');
  const inspect = (file, position) => execFileSync('python3', ['scripts/research/verify-combined-number-chart-native.py', file, String(position)], { encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'] });
  assert.equal(JSON.parse(inspect(candidate, 0)).cell, 7); assert.throws(() => inspect(candidate, 1));
  const original = readGeneratedZip(await readFile(candidate));
  const changes = [
    ['ppt/presentation.xml', xml => xml.replace('firstSlideNum="17"', 'firstSlideNum="18"')],
    ['ppt/slides/slide1.xml', xml => xml.replace('Combined chart and page number', 'wrong title')],
    ['ppt/charts/chart1.xml', xml => xml.replace('<c:v>7</c:v>', '<c:v>8</c:v>')],
    ['ppt/slides/slide2.xml', xml => xml.replace('Line values', 'unrelated')],
  ];
  for (const [index, [name, edit]] of changes.entries()) {
    const modified = new Map(original), before = modified.get(name).toString(), after = edit(before); assert.notEqual(after, before); modified.set(name, Buffer.from(after));
    const target = path.join(parent, `bad-${index}.pptx`); await writeFile(target, makeZip(modified));
    assert.throws(() => python(source, target));
    assert.throws(() => inspect(target, 0));
  }
  const altered = new Map(original), workbookName = 'ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx', workbook = readGeneratedZip(altered.get(workbookName));
  const sheet = workbook.get('xl/worksheets/sheet1.xml').toString(), badSheet = sheet.replace('<c r="B3"><v>7</v>', '<c r="B3"><v>8</v>'); assert.notEqual(badSheet, sheet);
  workbook.set('xl/worksheets/sheet1.xml', Buffer.from(badSheet)); altered.set(workbookName, makeZip(workbook));
  const bad = path.join(parent, 'bad-workbook.pptx'); await writeFile(bad, makeZip(altered)); assert.throws(() => python(source, bad)); assert.throws(() => inspect(bad, 0));
});

test('independent field/chart oracle rejects field identity, cache, chart and unrelated mutations', async t => {
  const parent = await mkdtemp(path.join(tmpdir(), 'pptkit-field-chart-evidence-')); t.after(() => rm(parent, { recursive: true, force: true }));
  const dir = path.join(parent, 'sample'); execFileSync(process.execPath, ['scripts/prepare-field-chart-native.mjs', dir]);
  const source = path.join(dir, 'source.pptx'), candidate = path.join(dir, 'candidate.pptx');
  const verify = target => execFileSync('python3', ['scripts/research/verify-field-chart.py', source, target], { encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'] });
  const result = JSON.parse(verify(candidate)); assert.equal(result.field_cache, 'STALE'); assert.equal(result.chart_value, 7);
  const original = readGeneratedZip(await readFile(candidate));
  const nativeCheck = (file, position) => execFileSync('python3', ['scripts/research/verify-field-chart-native.py', file, String(position)], { encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'] });
  const simulated = new Map(original), slideXml = simulated.get('ppt/slides/slide1.xml').toString();
  assert.ok(slideXml.includes('STALE')); simulated.set('ppt/slides/slide1.xml', Buffer.from(slideXml.replace('STALE', '17')));
  const updated = path.join(parent, 'simulated-native.pptx'); await writeFile(updated, makeZip(simulated));
  assert.equal(JSON.parse(nativeCheck(updated, 0)).field_cache, 17); assert.throws(() => nativeCheck(updated, 1));
  simulated.set('ppt/slides/slide1.xml', Buffer.from(slideXml.replace('STALE', '18'))); const badCache = path.join(parent, 'wrong-native-cache.pptx'); await writeFile(badCache, makeZip(simulated)); assert.throws(() => nativeCheck(badCache, 0));
  const changes = [
    ['ppt/slides/slide1.xml', xml => xml.replace('11111111-2222-3333-4444-555555555555', 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee')],
    ['ppt/slides/slide1.xml', xml => xml.replace('STALE', '999')],
    ['ppt/charts/chart1.xml', xml => xml.replace('<c:v>7</c:v>', '<c:v>8</c:v>')],
    ['ppt/slides/slide2.xml', xml => xml.replace('Line values', 'unrelated')],
  ];
  for (const [index, [name, edit]] of changes.entries()) {
    const modified = new Map(original), before = modified.get(name).toString(), after = edit(before); assert.notEqual(after, before); modified.set(name, Buffer.from(after));
    const target = path.join(parent, `bad-${index}.pptx`); await writeFile(target, makeZip(modified)); assert.throws(() => verify(target));
  }
  const altered = new Map(original), bookName = 'ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx', book = readGeneratedZip(altered.get(bookName));
  const xml = book.get('xl/worksheets/sheet1.xml').toString(), corrupt = xml.replace('<c r="B3"><v>7</v>', '<c r="B3"><v>8</v>'); assert.notEqual(corrupt, xml);
  book.set('xl/worksheets/sheet1.xml', Buffer.from(corrupt)); altered.set(bookName, makeZip(book));
  const badBook = path.join(parent, 'bad-book.pptx'); await writeFile(badBook, makeZip(altered)); assert.throws(() => verify(badBook));
});
