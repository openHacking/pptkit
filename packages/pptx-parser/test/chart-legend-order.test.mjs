import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import test from 'node:test';
import { planChartLegendOrder } from '../dist/import/chart-legend-order.js';
import { captureColumnLegendSources } from '../dist/import/chart-legend-source.js';
import { captureXmlTextSource } from '../dist/preservation/xml-text-source.js';
import { editXmlTextAttributes } from '../dist/preservation/replace-xml-relationship-id.js';
import { readGeneratedZip } from '../../../tests/support/read-generated-zip.mjs';

const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const entry = { elementIndex: 23, namespaceUri: C, position: 0, lexical: '00' };
test('legend remapping follows series identities rather than source series IDs or names', async () => {
  const edits = await planChartLegendOrder([entry], [17, 45, 2], [2, 17, 45], 100, async () => {});
  assert.deepEqual(edits.map(item => [item.elementIndex, item.expected, item.replacement]), [[23, '00', '1']]);
  assert.deepEqual(await planChartLegendOrder([entry], [17, 45], [17, 45], 100, async () => {}), []);
  for (const bad of [[17, 17], [17, 3], [17]]) await assert.rejects(planChartLegendOrder([entry], [17, 45], bad, 100, async () => {}), TypeError);
  for (const bad of [{ ...entry, position: 2, lexical: '2' }, { ...entry, lexical: '0x0' }, { ...entry, elementIndex: 0 }]) await assert.rejects(planChartLegendOrder([bad], [17, 45], [45, 17], 100, async () => {}), TypeError);
  await assert.rejects(planChartLegendOrder([entry, entry], [17, 45], [45, 17], 100, async () => {}), /duplicated/);
  await assert.rejects(planChartLegendOrder([entry], [17, 45], [45, 17], 1, async () => {}), RangeError);
  const cancel = new Error('cancel'); await assert.rejects(planChartLegendOrder([entry], [17, 45], [45, 17], 100, async () => { throw cancel; }), error => error === cancel);
  const source = [{ ...entry }], desired = [45, 17], pending = planChartLegendOrder(source, [17, 45], desired, 100, async () => {}); source[0].lexical = 'bad'; desired.reverse();
  assert.equal((await pending)[0].replacement, '1');
});

test('planned index edit reproduces the observed PowerPoint legend control', async () => {
  const read = async name => readGeneratedZip(await readFile(new URL(`../../../tests/fixtures/${name}.pptx`, import.meta.url))).get('ppt/charts/chart1.xml');
  const before = await read('chart-order-legend-before'), after = await read('chart-order-legend-after');
  const captured = await captureXmlTextSource(before); let index;
  try { index = captured.elements.find(item => item.local === 'idx' && captured.elements[item.parentElementIndex - 1]?.local === 'legendEntry').elementIndex; } finally { captured.dispose(); }
  const edits = await planChartLegendOrder([{ elementIndex: index, namespaceUri: C, position: 0, lexical: '0' }], [0, 1], [1, 0], 100, async () => {});
  const result = await editXmlTextAttributes(after, edits, { maxOutputBytes: 1_000_000, maxWorkUnits: 1_000_000, maxEdits: 10 });
  const text = after.toString(), expected = text.replace('<c:legendEntry><c:idx val="0"/>', '<c:legendEntry><c:idx val="1"/>');
  assert.notEqual(expected, text); assert.equal(new TextDecoder().decode(result.bytes), expected);
  assert.deepEqual(result.bytes, Uint8Array.from(await read('chart-order-legend-remapped')));
});

test('column legend capture proves raw ancestry, chart family and unique in-range positions', async () => {
  const bytes = readGeneratedZip(await readFile(new URL('../../../tests/fixtures/chart-order-legend-before.pptx', import.meta.url))).get('ppt/charts/chart1.xml');
  const capture = source => captureColumnLegendSources(source, {}, 100000, async () => {}), result = await capture(bytes);
  assert.deepEqual(result.physicalIds, [0, 1]); assert.equal(result.entries.length, 1); assert.equal(result.entries[0].position, 0);
  const edits = await planChartLegendOrder(result.entries, result.physicalIds, [1, 0], 100, async () => {}); assert.equal(edits[0].replacement, '1');
  const xml = bytes.toString();
  for (const altered of [
    xml.replace('<c:legendEntry><c:idx val="0"/>', '<c:legendEntry><c:idx val="2"/>'),
    xml.replace('<c:legendEntry><c:idx val="0"/>', '<c:legendEntry future="1"><c:idx val="0"/>'),
    xml.replace('<c:legendEntry><c:idx val="0"/>', '<c:legendEntry><c:idx val="0"/><c:idx val="1"/>'),
    xml.replace('<c:legendEntry><c:idx val="0"/>', '<c:legendEntry><c:idx val="0"/><c:extLst/>'),
    xml.replace('</c:chartSpace>', '<q:future xmlns:q="urn:future"/></c:chartSpace>'),
    xml.replace('<c:barDir val="col"/>', '<c:barDir val="bar"/>'),
    xml.replace('</c:ser>', '<c:trendline/></c:ser>'),
  ]) { assert.notEqual(altered, xml); await assert.rejects(capture(new TextEncoder().encode(altered)), TypeError); }
  await assert.rejects(captureColumnLegendSources(bytes, {}, 1, async () => {}), RangeError);
  const cancelled = new Error('cancel'); await assert.rejects(captureColumnLegendSources(bytes, {}, 100000, async () => { throw cancelled; }), error => error === cancelled);
});
