import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import test from 'node:test';
import { inspectColumnSeriesExtensions } from '../dist/import/chart-series-extension-source.js';
import { readGeneratedZip } from '../../../tests/support/read-generated-zip.mjs';

const chart = async name => readGeneratedZip(await readFile(new URL(`../../../tests/fixtures/${name}.pptx`, import.meta.url))).get('ppt/charts/chart1.xml');
const inspect = bytes => inspectColumnSeriesExtensions(bytes, {}, 1_000_000, async () => {});
test('recognizes observed series-owned extensions and stable identities after physical movement', async () => {
  const before = await inspect(await chart('chart-order-legend-powerpoint-resave')), after = await inspect(await chart('chart-order-powerpoint-extensions-moved'));
  assert.equal(before.length, 2); assert.deepEqual(before.map(item => item.seriesId), [1, 0]); assert.deepEqual(after, [...before].reverse());
  assert.notEqual(before[0].uniqueId, before[1].uniqueId);
});
test('rejects unknown extensions, duplicate identities, altered ownership and exhausted budgets', async () => {
  const bytes = await chart('chart-order-legend-powerpoint-resave'), xml = bytes.toString();
  const ids = [...xml.matchAll(/<c16:uniqueId val="([^"]+)"/gu)].map(match => match[1]); assert.equal(ids.length, 2);
  for (const changed of [
    xml.replace('{6F2FDCE9-48DA-4B69-8628-5D25D57E5C99}', '{00000000-0000-0000-0000-000000000000}'),
    xml.replace(ids[1], ids[0]),
    xml.replace('<a:srgbClr val="FFFFFF"/>', '<a:srgbClr val="FFFFFF"><a:alpha val="1"/></a:srgbClr>'),
    xml.replace('</c:chartSpace>', '<c:extLst/></c:chartSpace>'),
    xml.replace('</c:chartSpace>', '<c16:unknown xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart"/></c:chartSpace>'),
    xml.replace('<c:barDir val="col"/>', '<c:barDir val="bar"/>'),
  ]) { assert.notEqual(changed, xml); await assert.rejects(inspect(new TextEncoder().encode(changed)), TypeError); }
  await assert.rejects(inspectColumnSeriesExtensions(bytes, {}, 1, async () => {}), RangeError);
  const cancelled = new Error('cancelled'); await assert.rejects(inspectColumnSeriesExtensions(bytes, {}, 1_000_000, async () => { throw cancelled; }), error => error === cancelled);
});
