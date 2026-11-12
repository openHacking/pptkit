import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { spawnSync } from 'node:child_process';
import test from 'node:test';
import { importPptx } from '../packages/pptx-parser/dist/index.js';
import { replaceArchiveEntries } from '../packages/pptx-parser/dist/preservation/replace-archive-entries.js';
import { readGeneratedZip } from '../tests/support/read-generated-zip.mjs';

test('independent Python oracle verifies three category saves and rejects unrelated mutations', async () => {
  const directory = await mkdtemp(join(tmpdir(), 'pptkit-label-oracle-'));
  let bytes = await readFile(new URL('../tests/fixtures/chart-label-source.pptx', import.meta.url));
  const config = { projection: 'content', chartEditSaveProfile: 'existing-values-and-labels-v4' };
  const beforePath = join(directory, 'before.pptx'), afterPath = join(directory, 'after.pptx');
  const oracle = value => spawnSync('python3', ['scripts/research/verify-chart-label-save.py', beforePath, afterPath, JSON.stringify(value)], { encoding: 'utf8' });
  try {
    for (const value of [' & literal _x0041_ ', '中文 😀', 'last']) {
      const session = await importPptx(bytes, config);
      try {
        const slide = session.document.slides[0], chart = structuredClone(slide.elements[0]); chart.categories[1] = value;
        slide.removeElement(chart.id); slide.insertElement(0, chart);
        const plan = await session.prepareSave({ mode: 'edit' }); assert.equal(plan.status, 'ready', JSON.stringify(plan.report));
        const output = (await session.save(plan.plan)).bytes;
        await writeFile(beforePath, bytes); await writeFile(afterPath, output);
        const checked = oracle(value); assert.equal(checked.status, 0, checked.stderr);
        const parts = readGeneratedZip(output), chartXml = parts.get('ppt/charts/chart1.xml').toString();
        const altered = chartXml.replace('First _x005F_x005F_', 'Unexpected'); assert.notEqual(altered, chartXml);
        const bad = await replaceArchiveEntries(output, [{ name: 'ppt/charts/chart1.xml', bytes: new TextEncoder().encode(altered) }], { maxOutputBytes: 10_000_000 });
        await writeFile(afterPath, bad); assert.notEqual(oracle(value).status, 0);
        const diverged = chartXml.replace(/(<c:cat><c:strRef>[\s\S]*?<c:pt idx="1"><c:v>)[\s\S]*?(<\/c:v>)/, '$1diverged$2');
        assert.notEqual(diverged, chartXml);
        const badCache = await replaceArchiveEntries(output, [{ name: 'ppt/charts/chart1.xml', bytes: new TextEncoder().encode(diverged) }], { maxOutputBytes: 10_000_000 });
        await writeFile(afterPath, badCache); assert.notEqual(oracle(value).status, 0);
        bytes = output;
      } finally { session.dispose(); }
    }
  } finally { await rm(directory, { recursive: true, force: true }); }
});

test('independent oracle verifies xml:space insertion and default replacement without collateral edits', async () => {
  const directory = await mkdtemp(join(tmpdir(), 'pptkit-label-space-'));
  const beforePath = join(directory, 'before.pptx'), afterPath = join(directory, 'after.pptx');
  const original = await readFile(new URL('../tests/fixtures/chart-label-source.pptx', import.meta.url));
  const bookName = 'ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx', sheetName = 'xl/worksheets/sheet1.xml';
  const book = readGeneratedZip(original).get(bookName), sheet = readGeneratedZip(book).get(sheetName).toString();
  const target = '<t xml:space="preserve">literal _x005F_x0041_</t>'; assert.ok(sheet.includes(target));
  try {
    for (const attribute of ['', ' xml:space="default"']) {
      const changed = sheet.replace(target, `<t${attribute}>literal _x005F_x0041_</t>`);
      const workbook = await replaceArchiveEntries(book, [{ name: sheetName, bytes: new TextEncoder().encode(changed) }], { maxOutputBytes: 10_000_000 });
      const before = await replaceArchiveEntries(original, [{ name: bookName, bytes: workbook }], { maxOutputBytes: 10_000_000 });
      const session = await importPptx(before, { projection: 'content', chartEditSaveProfile: 'existing-values-and-labels-v4' });
      try {
        const slide = session.document.slides[0], chart = structuredClone(slide.elements[0]); chart.categories[1] = ' padded & _x0041_ '; slide.replaceElement(chart.id, chart);
        const plan = await session.prepareSave({ mode: 'edit' }); assert.equal(plan.status, 'ready', JSON.stringify(plan.report));
        const output = (await session.save(plan.plan)).bytes;
        await writeFile(beforePath, before); await writeFile(afterPath, output);
        const check = spawnSync('python3', ['scripts/research/verify-chart-label-save.py', beforePath, afterPath, JSON.stringify(chart.categories[1])], { encoding: 'utf8' });
        assert.equal(check.status, 0, check.stderr);
        const outputBook = readGeneratedZip(output).get(bookName), outputSheet = readGeneratedZip(outputBook).get(sheetName).toString();
        const broken = outputSheet.replace('<t xml:space="preserve"> Q1 </t>', '<t xml:space="default"> Q1 </t>'); assert.notEqual(broken, outputSheet);
        const badBook = await replaceArchiveEntries(outputBook, [{ name: sheetName, bytes: new TextEncoder().encode(broken) }], { maxOutputBytes: 10_000_000 });
        await writeFile(afterPath, await replaceArchiveEntries(output, [{ name: bookName, bytes: badBook }], { maxOutputBytes: 10_000_000 }));
        assert.notEqual(spawnSync('python3', ['scripts/research/verify-chart-label-save.py', beforePath, afterPath, JSON.stringify(chart.categories[1])]).status, 0);
      } finally { session.dispose(); }
    }
  } finally { await rm(directory, { recursive: true, force: true }); }
});

test('independent oracle verifies three series-name rounds and rejects chart or workbook corruption', async () => {
  const directory = await mkdtemp(join(tmpdir(), 'pptkit-series-name-'));
  const beforePath = join(directory, 'before.pptx'), afterPath = join(directory, 'after.pptx');
  let bytes = await readFile(new URL('../tests/fixtures/chart-label-source.pptx', import.meta.url));
  const config = { projection: 'content', chartEditSaveProfile: 'existing-values-and-series-names-v5' };
  const check = value => spawnSync('python3', ['scripts/research/verify-chart-series-name-save.py', beforePath, afterPath, JSON.stringify(value)], { encoding: 'utf8' });
  try {
    for (const value of [' & name _x0041_ ', '中文 😀', 'last']) {
      const session = await importPptx(bytes, config);
      try {
        const slide = session.document.slides[0], chart = structuredClone(slide.elements[0]); chart.series[0].name = value; slide.replaceElement(chart.id, chart);
        const plan = await session.prepareSave({ mode: 'edit' }); assert.equal(plan.status, 'ready', JSON.stringify(plan.report));
        const output = (await session.save(plan.plan)).bytes; await writeFile(beforePath, bytes); await writeFile(afterPath, output);
        const good = check(value); assert.equal(good.status, 0, good.stderr);
        const parts = readGeneratedZip(output), chartName = 'ppt/charts/chart1.xml', bookName = 'ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx', sheetName = 'xl/worksheets/sheet1.xml';
        const chartXml = parts.get(chartName).toString(), brokenChart = chartXml.replace('Second &amp; series', 'other series'); assert.notEqual(brokenChart, chartXml);
        await writeFile(afterPath, await replaceArchiveEntries(output, [{ name: chartName, bytes: new TextEncoder().encode(brokenChart) }], { maxOutputBytes: 10_000_000 })); assert.notEqual(check(value).status, 0);
        const book = parts.get(bookName), sheet = readGeneratedZip(book).get(sheetName).toString(), brokenSheet = sheet.replace(/(<c r="B1"[^>]*><is><t[^>]*>)[\s\S]*?(<\/t>)/, '$1wrong$2'); assert.notEqual(brokenSheet, sheet);
        const badBook = await replaceArchiveEntries(book, [{ name: sheetName, bytes: new TextEncoder().encode(brokenSheet) }], { maxOutputBytes: 10_000_000 });
        await writeFile(afterPath, await replaceArchiveEntries(output, [{ name: bookName, bytes: badBook }], { maxOutputBytes: 10_000_000 })); assert.notEqual(check(value).status, 0);
        bytes = output;
      } finally { session.dispose(); }
    }
  } finally { await rm(directory, { recursive: true, force: true }); }
});
