import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import test from 'node:test';
import { importPptx } from '../dist/index.js';
import { readGeneratedZip } from '../../../tests/support/read-generated-zip.mjs';
import { replaceArchiveEntries } from '../dist/preservation/replace-archive-entries.js';

const options = { projection: 'content', chartEditSaveProfile: 'existing-series-order-v6' };
const source = () => readFile(new URL('../../../tests/fixtures/chart-label-source.pptx', import.meta.url));
function edit(session, mutate) {
  const slide = session.document.slides[0], chart = structuredClone(slide.elements.find(element => element.type === 'chart'));
  mutate(chart); slide.replaceElement(chart.id, chart);
}

test('series-order requests remain blocked after the PowerPoint counterexample', async () => {
  const session = await importPptx(await source(), options);
  try {
    const binding = session.report.chartSeriesOrders[0]; assert.equal(binding.status, 'bound');
    edit(session, chart => chart.series.reverse());
    const result = await session.prepareSave({ mode: 'edit', chartSeriesOrders: [{ slideId: binding.slideId, elementId: binding.elementId, identities: binding.series.map(item => item.sourceId).reverse() }] });
    assert.equal(result.status, 'blocked'); assert.equal(result.report.diagnostics[0].code, 'save-chart-order-native-unsupported');
  } finally { session.dispose(); }
});

test('series-order requests remain blocked with signatures or package-wide extensions', async () => {
  const bytes = await source(), parts = readGeneratedZip(bytes);
  const relName = 'ppt/_rels/presentation.xml.rels', rels = parts.get(relName).toString();
  const signed = rels.replace('<Relationships ', '<Relationships xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future" mc:Ignorable="x" ').replace('</Relationships>', '<mc:AlternateContent><mc:Choice Requires="x"><Relationship Id="hidden" Type="http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature" Target="missing.bin"/></mc:Choice><mc:Fallback/></mc:AlternateContent></Relationships>');
  const presentationName = 'ppt/presentation.xml', presentation = parts.get(presentationName).toString();
  for (const [name, xml] of [[relName, signed], [presentationName, presentation.replace('</p:presentation>', '<p:extLst/></p:presentation>')]]) {
    const input = await replaceArchiveEntries(bytes, [{ name, bytes: new TextEncoder().encode(xml) }], { maxOutputBytes: 10_000_000 });
    const session = await importPptx(input, options);
    try {
      const binding = session.report.chartSeriesOrders[0];
      const identities = binding.status === 'bound' ? binding.series.map(item => item.sourceId).reverse() : [1, 0];
      edit(session, chart => chart.series.reverse());
      const result = await session.prepareSave({ mode: 'edit', chartSeriesOrders: [{ slideId: binding.slideId, elementId: binding.elementId, identities }] });
      assert.equal(result.status, 'blocked'); assert.equal(result.report.diagnostics[0].code, 'save-chart-order-native-unsupported');
    } finally { session.dispose(); }
  }
});

test('series-order session requires explicit source identity and snapshots requests', async () => {
  const session = await importPptx(await source(), options);
  try {
    const binding = session.report.chartSeriesOrders[0], ids = binding.series.map(item => item.sourceId).reverse();
    edit(session, chart => chart.series.reverse());
    assert.equal((await session.prepareSave({ mode: 'edit' })).status, 'blocked');
    const request = { slideId: binding.slideId, elementId: binding.elementId, identities: ids };
    const pending = session.prepareSave({ mode: 'edit', chartSeriesOrders: [request] }); ids.reverse();
    const prepared = await pending; assert.equal(prepared.status, 'blocked');
    assert.equal((await session.prepareSave({ mode: 'edit', chartSeriesOrders: [request] })).status, 'blocked');
    await assert.rejects(session.prepareSave({ mode: 'no-op', chartSeriesOrders: [request] }), TypeError);
  } finally { session.dispose(); }
});

test('PowerPoint resave rebinds scalar sources while order remains blocked', async () => {
  const bytes = await readFile(new URL('../../../tests/fixtures/chart-order-legend-powerpoint-resave.pptx', import.meta.url));
  const session = await importPptx(bytes, options);
  try {
    assert.ok(session.report.chartValues.every(item => item.status === 'bound'));
    assert.ok(session.report.chartSeriesNames.every(item => item.status === 'bound'));
    assert.equal(session.report.chartSeriesOrders[0].status, 'unmapped');
    edit(session, chart => { chart.series[0].name = 'After resave'; chart.series[0].values[0] = 4.5; chart.categories[1] = 'After category'; });
    const plan = await session.prepareSave({ mode: 'edit' }); assert.equal(plan.status, 'ready', JSON.stringify(plan.report));
    const reopened = await importPptx((await session.save(plan.plan)).bytes, options);
    try {
      const chart = reopened.document.slides[0].elements.find(item => item.type === 'chart');
      assert.equal(chart.series[0].name, 'After resave'); assert.equal(chart.series[0].values[0], 4.5); assert.equal(chart.categories[1], 'After category');
      assert.equal(reopened.report.chartSeriesOrders[0].status, 'unmapped');
    } finally { reopened.dispose(); }
    plan.plan.dispose();
  } finally { session.dispose(); }
});
