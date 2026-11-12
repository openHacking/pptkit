import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import test from 'node:test';
import { readOpcSource } from '../dist/opc/read-package.js';
import { inspectColumnSeriesExtensions } from '../dist/import/chart-series-extension-source.js';
import { inspectChartExtensionReferences } from '../dist/import/chart-extension-references.js';
import { replaceArchiveEntries } from '../dist/preservation/replace-archive-entries.js';
import { readGeneratedZip } from '../../../tests/support/read-generated-zip.mjs';

const fixture = () => readFile(new URL('../../../tests/fixtures/chart-order-legend-powerpoint-resave.pptx', import.meta.url));
const inspect = (source, ids, work = 10_000_000, checkpoint = async () => {}) => inspectChartExtensionReferences(source, '/ppt/charts/chart1.xml', ids, work, checkpoint);
async function prepare(bytes, retain = true) {
  const parts = readGeneratedZip(bytes), ids = (await inspectColumnSeriesExtensions(parts.get('ppt/charts/chart1.xml'), {}, 1_000_000, async () => {})).map(item => item.uniqueId);
  return { source: await readOpcSource(bytes, {}, retain, true), ids };
}
test('finds no other literal XML references but reports opaque workbook coverage', async () => {
  const { source, ids } = await prepare(await fixture()), result = await inspect(source, ids);
  assert.deepEqual(result.references, []); assert.ok(result.opaqueParts.some(part => part.endsWith('.xlsx'))); assert.ok(result.workUnits > 0);
  await assert.rejects(inspect(source, ids, result.workUnits - 1), RangeError);
  assert.equal((await inspect(source, ids, result.workUnits)).workUnits, result.workUnits);
  const missing = await prepare(await fixture(), false); await assert.rejects(inspect(missing.source, ids), /every original XML/);
});

test('detects references in XML attributes, comments and entity-normalized text outside the chart', async () => {
  const bytes = await fixture(), parts = readGeneratedZip(bytes), name = 'ppt/slides/slide1.xml', xml = parts.get(name).toString();
  const { ids } = await prepare(bytes), guid = ids[0], escaped = guid.replace('0', '&#48;');
  const altered = xml.replace('</p:sld>', `<!-- ${guid.toLowerCase()} --><?ref ${guid}?><p:extLst><p:ext uri="${guid}">${escaped}</p:ext></p:extLst><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:q="urn:future"><mc:Choice Requires="q"><q:ref val="${guid}"/></mc:Choice><mc:Fallback/></mc:AlternateContent></p:sld>`);
  assert.notEqual(altered, xml);
  const input = await replaceArchiveEntries(bytes, [{ name, bytes: new TextEncoder().encode(altered) }], { maxOutputBytes: 10_000_000 });
  const result = await inspect((await prepare(input)).source, ids), hits = result.references.filter(item => item.part === '/ppt/slides/slide1.xml');
  assert.deepEqual(hits.map(item => item.kind).sort(), ['attribute', 'attribute', 'comment', 'processing-instruction', 'text']); assert.ok(hits.every(item => item.uniqueId === guid));
  const cancelled = new Error('cancel'); await assert.rejects(inspect((await prepare(input)).source, ids, 10_000_000, async () => { throw cancelled; }), error => error === cancelled);
});

test('rejects duplicate declarations and preserves opaque coverage on empty XML-reference results', async () => {
  const bytes = await fixture(), parts = readGeneratedZip(bytes), name = 'ppt/charts/chart1.xml', xml = parts.get(name).toString(), { ids } = await prepare(bytes);
  const changed = xml.replace(ids[1], ids[0]); assert.notEqual(changed, xml);
  const input = await replaceArchiveEntries(bytes, [{ name, bytes: new TextEncoder().encode(changed) }], { maxOutputBytes: 10_000_000 });
  const raw = await readOpcSource(input, {}, true, true);
  await assert.rejects(inspect(raw, [ids[0]]), /exactly one chart declaration/);
  await assert.rejects(inspect(raw, [ids[0], ids[0]]), TypeError);
});
