import test from 'node:test';
import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdtemp, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { promisify } from 'node:util';
import { fileURLToPath } from 'node:url';
import { renderDashboard } from '../examples/report-dashboard/dashboard.mjs';
import { renderDiagramTemplate } from '../examples/report-dashboard/diagram-template.mjs';
import { importPptx } from '../packages/pptx-parser/dist/index.js';
import { readGeneratedZip } from '../tests/support/read-generated-zip.mjs';

test('report example exports editable data and a matching slide preview', async () => {
  const { diagnostics, pptx, preview } = await renderDashboard();
  assert.equal(diagnostics.filter(item => item.severity === 'error').length, 0);
  assert.equal(pptx.slideCount, 1);
  assert.equal(preview.slides.length, 1);
  assert.match(preview.slides[0].svg, /Quarterly operations report/);
  assert.match(preview.slides[0].svg, /Retention/);
  const parts = readGeneratedZip(pptx.bytes);
  assert.ok([...parts.keys()].some(name => name.endsWith('.xlsx')));
  const session = await importPptx(pptx.bytes, { projection: 'content' });
  try {
    const elements = session.document.slides[0].elements;
    const chart = elements.find(element => element.type === 'chart');
    const table = elements.find(element => element.type === 'table');
    assert.deepEqual(chart.categories, ['Q1', 'Q2', 'Q3', 'Q4']);
    assert.deepEqual(chart.series[0].values, [110, 125, 141, 158]);
    assert.equal(table.rows.length, 3);
    assert.equal(table.rows[2].cells.length, 2);
  } finally { session.dispose(); }
});

test('example command writes files and refuses to overwrite them', async t => {
  const dir = await mkdtemp(path.join(tmpdir(), 'pptkit-report-example-'));
  t.after(() => rm(dir, { recursive: true, force: true }));
  const entry = fileURLToPath(new URL('../examples/report-dashboard/dashboard.mjs', import.meta.url));
  const run = promisify(execFile);
  const result = await run(process.execPath, [entry], { cwd: dir });
  const report = JSON.parse(result.stdout);
  assert.ok(Array.isArray(report.exportWarnings));
  assert.ok(Array.isArray(report.previewWarnings));
  const pptx = await readFile(path.join(dir, 'report-dashboard.pptx'));
  assert.match(await readFile(path.join(dir, 'report-dashboard.svg'), 'utf8'), /Quarterly operations report/);
  await assert.rejects(run(process.execPath, [entry], { cwd: dir }));
  assert.deepEqual(await readFile(path.join(dir, 'report-dashboard.pptx')), pptx);
});

test('diagram template shares its layout and exports connected native shapes', async () => {
  const { diagnostics, pptx, preview } = await renderDiagramTemplate();
  assert.equal(diagnostics.filter(item => item.severity === 'error').length, 0);
  assert.equal(pptx.slideCount, 2);
  assert.equal(preview.slides.length, 2);
  assert.match(preview.slides[0].svg, /Plan/);
  assert.match(preview.slides[1].svg, /Next steps/);
  const parts = readGeneratedZip(pptx.bytes);
  const first = parts.get('ppt/slides/slide1.xml').toString();
  assert.match(first, /<p:cxnSp>/);
  assert.match(first, /<a:stCxn id="\d+" idx="3"\/>/);
  assert.match(first, /<a:endCxn id="\d+" idx="1"\/>/);
  const layoutTargets = [1, 2].map(index => {
    const rels = parts.get(`ppt/slides/_rels/slide${index}.xml.rels`).toString();
    return rels.match(/Type="[^"]*\/slideLayout" Target="([^"]+)"/)?.[1];
  });
  assert.ok(layoutTargets[0]);
  assert.equal(layoutTargets[0], layoutTargets[1]);
  const session = await importPptx(pptx.bytes, { projection: 'content' });
  try {
    assert.equal(session.document.slides.length, 2);
    assert.equal(session.document.slides[0].elements.filter(element => element.type === 'shape').length, 2);
    assert.equal(session.document.slides[0].elements.filter(element => element.type === 'connector').length, 1);
  } finally { session.dispose(); }
});

test('diagram template command creates a file without overwriting it', async t => {
  const dir = await mkdtemp(path.join(tmpdir(), 'pptkit-diagram-example-'));
  t.after(() => rm(dir, { recursive: true, force: true }));
  const entry = fileURLToPath(new URL('../examples/report-dashboard/diagram-template.mjs', import.meta.url));
  const run = promisify(execFile);
  const result = await run(process.execPath, [entry], { cwd: dir });
  assert.ok(Array.isArray(JSON.parse(result.stdout).exportWarnings));
  const original = await readFile(path.join(dir, 'diagram-template.pptx'));
  await assert.rejects(run(process.execPath, [entry], { cwd: dir }));
  assert.deepEqual(await readFile(path.join(dir, 'diagram-template.pptx')), original);
});
