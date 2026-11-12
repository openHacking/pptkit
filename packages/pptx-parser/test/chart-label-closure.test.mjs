import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import test from 'node:test';
import { DEFAULT_SESSION_LIMITS } from '../dist/index.js';
import { chartLabelClosure } from '../dist/import/chart-label-closure.js';
import { chartReference } from '../dist/import/chart-data.js';
import { chartWorkbooks, workbookRange } from '../dist/import/chart-workbook.js';
import { contentReader } from '../dist/import/content-reader.js';
import { readXml } from '../dist/xml/read-xml.js';
import { readGeneratedZip } from '../../../tests/support/read-generated-zip.mjs';
import { structureContext } from '../dist/presentation/structure-context.js';
import { readStructureFromContext } from '../dist/presentation/read-structure.js';
import { contentCharts } from '../dist/import/content-chart.js';
import { replaceArchiveEntries } from '../dist/preservation/replace-archive-entries.js';
import { readOpcSource } from '../dist/opc/read-package.js';
import { prepareChartValues } from '../dist/import/prepare-chart-values.js';
import { analyzeTextEdits } from '../dist/import/analyze-text-edits.js';
import { projectStructure } from '../dist/import/project-structure.js';
import { projectContent } from '../dist/import/project-content.js';
import { snapshot } from '../dist/import/state.js';
import { prepareTextCandidate } from '../dist/import/prepare-text-candidate.js';
import { analyzeChartOrders, captureChartSeriesOrderRequests } from '../dist/import/analyze-chart-order.js';
import { inspectClosure } from '../dist/preservation/inspect-closure.js';
import { verifyChartOrderRelationshipClosure } from '../dist/import/chart-order-closure.js';
import { planChartSeriesOrder, verifyChartSeriesOrderRebound } from '../dist/import/chart-series-order.js';
import { editXmlTextAttributes } from '../dist/preservation/replace-xml-relationship-id.js';
import { chartBytes, chartXml, series, workbook } from './chart-fixture.mjs';

const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const cellKey = cell => JSON.stringify([cell.workbookPart, cell.worksheetPart, cell.reference]);
const keys = ['category:0', 'category:1', 'category:2'];
test('chart order requests capture explicit identities and reject ambiguous input', () => {
  const request = { slideId: 'slide', elementId: 'chart', identities: [1, 0] };
  const captured = captureChartSeriesOrderRequests([request], 3, 100);
  request.identities.reverse(); request.slideId = 'changed';
  assert.deepEqual(captured, [{ slideId: 'slide', elementId: 'chart', identities: [1, 0] }]);
  assert.deepEqual(captureChartSeriesOrderRequests(undefined, 3, 100), []);
  for (const invalid of [null, [request, request], [{ ...request, identities: [1, 1] }], [{ ...request, identities: [-0] }], [{ ...request, identities: [4294967296] }], [{ ...request, identities: [] }], [{ ...request, extra: true }], new Array(1)]) {
    assert.throws(() => captureChartSeriesOrderRequests(invalid, 3, 100), TypeError);
  }
  let invoked = false;
  const accessor = { elementId: 'chart', identities: [0], get slideId() { invoked = true; return 'slide'; } };
  assert.throws(() => captureChartSeriesOrderRequests([accessor], 3, 100), TypeError); assert.equal(invoked, false);
  const array = [0]; Object.defineProperty(array, '0', { get() { invoked = true; return 0; } });
  assert.throws(() => captureChartSeriesOrderRequests([{ ...request, identities: array }], 3, 100), TypeError); assert.equal(invoked, false);
  assert.throws(() => captureChartSeriesOrderRequests([request], 3, 2), /work budget/);
  assert.throws(() => captureChartSeriesOrderRequests([request, { ...request, elementId: 'other' }], 1, 100), /maxSaveChanges/);
});
async function observations() {
  const fixture = await readFile(new URL('../../../tests/fixtures/chart-label-source.pptx', import.meta.url)), parts = readGeneratedZip(fixture);
  const chart = (await readXml(parts.get('ppt/charts/chart1.xml'), {}, true)).root;
  const indices = new Map(), todo = [chart]; let index = 0;
  while (todo.length) { const node = todo.pop(); indices.set(node, ++index); const children = node.children.filter(n => n.kind === 'element'); todo.push(...children.reverse()); }
  const book = await chartWorkbooks({}, DEFAULT_SESSION_LIMITS, () => {}, async () => {}, true)('/book.xlsx', parts.get('ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx'));
  const r = contentReader('/chart.xml', P, A), nodes = [...indices.keys()].filter(n => n.namespaceUri === C && n.local === 'cat');
  const uses = [], counts = new Map();
  for (const node of nodes) {
    const ref = await chartReference(node, false, C, r, async () => {}, () => {}, true);
    const range = await workbookRange(book, ref.formula, false, () => {}, async () => {}, ref.values.length);
    for (const [i, point] of ref.points.entries()) {
      const entry = range.cells[i], source = entry.cell.stringSource;
      assert.ok(source); const cell = { workbookPart: '/book.xlsx', worksheetPart: source.part, reference: entry.reference, elementIndex: source.elementIndex, lexical: source.lexical };
      uses.push({ logicalKey: `category:${i}`, projected: ref.values[i], cache: { part: '/chart.xml', elementIndex: indices.get(point.node), lexical: point.lexical }, cell });
      counts.set(cellKey(cell), (counts.get(cellKey(cell)) ?? 0) + 1);
    }
  }
  return { uses, counts };
}

test('closure groups shared categories from actual chart and workbook sources', async () => {
  const { uses, counts } = await observations();
  const result = await chartLabelClosure(keys, uses, counts, true, 100, async () => {});
  assert.deepEqual(result.map(item => [item.logicalKey, item.status, item.uses.length]), [['category:0', 'bound', 2], ['category:1', 'bound', 2], ['category:2', 'bound', 2]]);
  assert.equal(result[1].uses[0].cache.lexical, 'literal _x005F_x0041_');
});

test('closure blocks unknown references, conflicting owners, duplicate caches and inconsistent strings', async () => {
  const { uses, counts } = await observations();
  async function status(items, observed = counts, complete = true) { return (await chartLabelClosure(keys, items, observed, complete, 100, async () => {})).find(item => item.logicalKey === 'category:1'); }
  assert.equal((await status(uses, counts, false)).status, 'unmapped');
  const extra = new Map(counts); extra.set(cellKey(uses[1].cell), 3); assert.match((await status(uses, extra)).reason, /unaccounted/);
  const conflicting = structuredClone(uses); conflicting[4].logicalKey = 'category:0'; assert.match((await status(conflicting)).reason, /conflicting/);
  const duplicate = structuredClone(uses); duplicate[4].cache = duplicate[1].cache; assert.match((await status(duplicate)).reason, /assigned/);
  const altered = structuredClone(uses); altered[4].cache.lexical = 'other'; assert.match((await status(altered)).reason, /disagree/);
  const source = structuredClone(uses); source[4].cell.elementIndex++; assert.match((await status(source)).reason, /physical sources/);
  assert.match((await status(uses.filter(use => use.logicalKey !== 'category:1'))).reason, /No source/);
  await assert.rejects(chartLabelClosure(keys, uses, counts, true, 3, async () => {}), RangeError);
  await assert.rejects(chartLabelClosure([...keys, keys[0]], uses, counts, true, 100, async () => {}), TypeError);
  const abort = new Error('cancelled'); await assert.rejects(chartLabelClosure(keys, uses, counts, true, 100, async () => { throw abort; }), error => error === abort);
});

async function inspectThroughImporter(input, capture = true, limits = DEFAULT_SESSION_LIMITS, breakAncestor, captureNames = false, captureOrder = false) {
  const context = await structureContext(input, undefined, true), inspection = await readStructureFromContext(context);
  if (breakAncestor) {
    const sourceElement = context.sourceElement;
    context.sourceElement = async (part, node) => {
      const source = await sourceElement(part, node);
      return part.endsWith('/chart1.xml') && node.local === breakAncestor ? { ...source, parent: undefined } : source;
    };
  }
  const part = inspection.slides[0].part, root = await context.root(part, 'sld');
  const todo = [root]; let graphic;
  while (todo.length) { const node = todo.pop(); if (node.local === 'graphicData') { graphic = node; break; } todo.push(...node.children.filter(child => child.kind === 'element')); }
  assert.ok(graphic);
  const read = contentCharts(context, limits, {}, inspection.package.relationships, async () => {}, true, false, true, capture, captureNames, captureOrder);
  return read(graphic, { x: 40, y: 40, width: 500, height: 300 }, contentReader(part, context.p, context.a), '2', () => {});
}

test('private source-order planning changes only order attributes and reprojects series identities', async () => {
  const source = await readFile(new URL('../../../tests/fixtures/chart-label-source.pptx', import.meta.url));
  const inspect = bytes => inspectThroughImporter(bytes, false, DEFAULT_SESSION_LIMITS, undefined, false, true);
  const before = await inspect(source), candidate = before.seriesOrderCandidate;
  assert.equal(candidate.complete, true); assert.equal(candidate.localDependencyReason, undefined); assert.deepEqual(candidate.sources.map(item => [item.identity, item.order]), [[0, 0], [1, 1]]);
  const edits = await planChartSeriesOrder(candidate.sources, [1, 0], candidate.complete, 100, async () => {}); assert.equal(edits.length, 2);
  const parts = readGeneratedZip(source), chartName = 'ppt/charts/chart1.xml';
  const changed = await editXmlTextAttributes(parts.get(chartName), edits, { maxOutputBytes: 1_000_000, maxEdits: 10, maxWorkUnits: 1_000_000 });
  const output = await replaceArchiveEntries(source, [{ name: chartName, bytes: changed.bytes }], { maxOutputBytes: 10_000_000 });
  const after = await inspect(output); assert.deepEqual(after.fields.series.map(series => series.name), ['Second & series', 'First _x005F_']);
  assert.deepEqual(after.fields.series.map(series => series.values), [[4, 5, 6], [1, 2, 3]]);
  assert.deepEqual(after.seriesOrderCandidate.sources.map(item => item.identity), [1, 0]);
  const rebound = after.seriesOrderCandidate;
  await verifyChartSeriesOrderRebound(candidate.sources, [1, 0], rebound.sources, true, rebound.complete, 100, async () => {});
  const wrong = structuredClone(rebound.sources); wrong[0].elementIndex += 1;
  await assert.rejects(verifyChartSeriesOrderRebound(candidate.sources, [1, 0], wrong, true, true, 100, async () => {}), /identity|location/);
  const wrongRank = structuredClone(rebound.sources); [wrongRank[0].order, wrongRank[1].order] = [wrongRank[1].order, wrongRank[0].order]; [wrongRank[0].orderLexical, wrongRank[1].orderLexical] = [wrongRank[1].orderLexical, wrongRank[0].orderLexical];
  await assert.rejects(verifyChartSeriesOrderRebound(candidate.sources, [1, 0], wrongRank, true, true, 100, async () => {}), /order/);
  await assert.rejects(verifyChartSeriesOrderRebound(candidate.sources, [1, 0], rebound.sources, true, false, 100, async () => {}), /complete/);
  await assert.rejects(verifyChartSeriesOrderRebound(candidate.sources, [1, 0], rebound.sources, true, true, 1, async () => {}), RangeError);
  for (const [name, bytes] of parts) if (name !== chartName) assert.deepEqual(readGeneratedZip(output).get(name), bytes, name);
  assert.deepEqual(await planChartSeriesOrder(candidate.sources, [0, 1], true, 100, async () => {}), []);
  for (const bad of [[0, 0], [0, 2], [0]]) await assert.rejects(planChartSeriesOrder(candidate.sources, bad, true, 100, async () => {}), TypeError);
  await assert.rejects(planChartSeriesOrder(candidate.sources, [1, 0], false, 100, async () => {}), TypeError);
  await assert.rejects(planChartSeriesOrder(candidate.sources, [1, 0], true, 1, async () => {}), RangeError);
  const abort = new Error('cancelled'); await assert.rejects(planChartSeriesOrder(candidate.sources, [1, 0], true, 100, async () => { throw abort; }), error => error === abort);
  const otherFamily = structuredClone(candidate.sources); otherFamily[1].family = 'lineChart';
  await assert.rejects(planChartSeriesOrder(otherFamily, [1, 0], true, 100, async () => {}), /boundaries/);
  const invalid = structuredClone(candidate.sources); invalid[1].elementIndex = invalid[0].elementIndex;
  await assert.rejects(planChartSeriesOrder(invalid, [1, 0], true, 100, async () => {}), /duplicated/);
  const mutable = structuredClone(candidate.sources), order = [1, 0], pending = planChartSeriesOrder(mutable, order, true, 100, async () => {});
  mutable[0].orderLexical = 'changed'; order.reverse(); assert.deepEqual(await pending, edits);
});

test('order inspection flags chart extensions and slide timing without claiming save authority', async () => {
  const bytes = await readFile(new URL('../../../tests/fixtures/chart-label-source.pptx', import.meta.url)), parts = readGeneratedZip(bytes);
  const inspect = input => inspectThroughImporter(input, false, DEFAULT_SESSION_LIMITS, undefined, false, true);
  const chartName = 'ppt/charts/chart1.xml', chart = parts.get(chartName).toString();
  const changedChart = chart.replace('</c:chartSpace>', '<c:extLst/></c:chartSpace>'); assert.notEqual(changedChart, chart);
  const withExtension = await replaceArchiveEntries(bytes, [{ name: chartName, bytes: new TextEncoder().encode(changedChart) }], { maxOutputBytes: 10_000_000 });
  const extension = (await inspect(withExtension)).seriesOrderCandidate;
  assert.equal(extension.complete, false); assert.match(extension.localDependencyReason, /extensions/);
  const slideName = 'ppt/slides/slide1.xml', slide = parts.get(slideName).toString();
  for (const extra of ['<p:timing/>', '<p:extLst/>']) {
    const changed = slide.replace('</p:sld>', `${extra}</p:sld>`); assert.notEqual(changed, slide);
    const input = await replaceArchiveEntries(bytes, [{ name: slideName, bytes: new TextEncoder().encode(changed) }], { maxOutputBytes: 10_000_000 });
    assert.match((await inspect(input)).seriesOrderCandidate.localDependencyReason, /timing or extensions/);
  }
  const foreign = chart.replace('</c:chartSpace>', '<q:future xmlns:q="urn:example"/></c:chartSpace>');
  const foreignInput = await replaceArchiveEntries(bytes, [{ name: chartName, bytes: new TextEncoder().encode(foreign) }], { maxOutputBytes: 10_000_000 });
  assert.match((await inspect(foreignInput)).seriesOrderCandidate.localDependencyReason, /foreign elements/);
  const relName = 'ppt/charts/_rels/chart1.xml.rels', rels = parts.get(relName).toString();
  const extraRelation = rels.replace('</Relationships>', '<Relationship Id="rIdExtra" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/" TargetMode="External"/></Relationships>');
  const relationInput = await replaceArchiveEntries(bytes, [{ name: relName, bytes: new TextEncoder().encode(extraRelation) }], { maxOutputBytes: 10_000_000 });
  assert.match((await inspect(relationInput)).seriesOrderCandidate.localDependencyReason, /relationships/);
});

test('order inspection conservatively rejects package-wide extensions and timing', async () => {
  const bytes = await readFile(new URL('../../../tests/fixtures/chart-label-source.pptx', import.meta.url)), parts = readGeneratedZip(bytes);
  const name = 'ppt/presentation.xml', xml = parts.get(name).toString();
  for (const extra of ['<p:extLst/>', '<p:timing/>']) {
    const changed = xml.replace('</p:presentation>', `${extra}</p:presentation>`); assert.notEqual(changed, xml);
    const input = await replaceArchiveEntries(bytes, [{ name, bytes: new TextEncoder().encode(changed) }], { maxOutputBytes: 10_000_000 });
    const result = await inspectThroughImporter(input, false, DEFAULT_SESSION_LIMITS, undefined, false, true);
    assert.equal(result.seriesOrderCandidate.complete, false);
    assert.match(result.seriesOrderCandidate.localDependencyReason, /Package-wide timing or extensions/);
  }
});

test('order candidates reject chart-level indexed dependencies without disabling scalar bindings', async () => {
  const bytes = await readFile(new URL('../../../tests/fixtures/chart-label-source.pptx', import.meta.url)), name = 'ppt/charts/chart1.xml';
  const xml = readGeneratedZip(bytes).get(name).toString();
  const cases = [
    xml.replace('<c:legend><c:legendPos val="r"/>', '<c:legend><c:legendPos val="r"/><c:legendEntry><c:idx val="0"/><c:delete val="1"/></c:legendEntry>'),
    xml.replace('</c:barChart>', '<c:dLbls><c:dLbl><c:idx val="0"/></c:dLbl></c:dLbls></c:barChart>'),
    xml.replace('</c:chart>', '<c:pivotFmts><c:pivotFmt><c:idx val="0"/></c:pivotFmt></c:pivotFmts></c:chart>'),
  ];
  for (const changed of cases) {
    assert.notEqual(changed, xml);
    const input = await replaceArchiveEntries(bytes, [{ name, bytes: new TextEncoder().encode(changed) }], { maxOutputBytes: 10_000_000 });
    const result = await inspectThroughImporter(input, true, DEFAULT_SESSION_LIMITS, undefined, true, true);
    assert.equal(result.seriesOrderCandidate.complete, false); assert.match(result.seriesOrderCandidate.localDependencyReason, /Indexed chart-level/);
    assert.ok(result.valueBindings.some(binding => binding.status === 'bound'));
  }
});

test('order source inspection rejects opaque order metadata and retains integer lexemes', async () => {
  const bytes = await readFile(new URL('../../../tests/fixtures/chart-label-source.pptx', import.meta.url));
  const name = 'ppt/charts/chart1.xml', xml = readGeneratedZip(bytes).get(name).toString();
  const inspect = value => inspectThroughImporter(value, false, DEFAULT_SESSION_LIMITS, undefined, false, true);
  const source = xml.replace('<c:order val="0"/>', '<c:order val="002"/>').replace('<c:order val="1"/>', '<c:order val="007"/>'); assert.notEqual(source, xml);
  const input = await replaceArchiveEntries(bytes, [{ name, bytes: new TextEncoder().encode(source) }], { maxOutputBytes: 10_000_000 });
  const before = await inspect(input), candidate = before.seriesOrderCandidate;
  assert.equal(candidate.complete, true);
  const edits = await planChartSeriesOrder(candidate.sources, [1, 0], true, 100, async () => {});
  assert.deepEqual(edits.map(edit => [edit.expected, edit.replacement]), [['007', '002'], ['002', '007']]);
  const opaque = xml.replace('<c:order val="0"/>', '<c:order val="0" future="1"/>'); assert.notEqual(opaque, xml);
  const invalid = await replaceArchiveEntries(bytes, [{ name, bytes: new TextEncoder().encode(opaque) }], { maxOutputBytes: 10_000_000 });
  assert.equal((await inspect(invalid)).seriesOrderCandidate.complete, false);
  const unknown = xml.replace('</c:chartSpace>', '<c:extLst><c:ext uri="future"><c:f>Sheet1!$A$3</c:f></c:ext></c:extLst></c:chartSpace>');
  const uncertain = await replaceArchiveEntries(bytes, [{ name, bytes: new TextEncoder().encode(unknown) }], { maxOutputBytes: 10_000_000 });
  assert.equal((await inspect(uncertain)).seriesOrderCandidate.complete, false);
});

test('private chart staging combines order, cache and workbook edits without overwriting them', async () => {
  const bytes = await readFile(new URL('../../../tests/fixtures/chart-label-source.pptx', import.meta.url));
  const result = await inspectThroughImporter(bytes, true, DEFAULT_SESSION_LIMITS, undefined, false, true), source = await readOpcSource(bytes, {}, false, true);
  const order = { part: '/ppt/charts/chart1.xml', sources: result.seriesOrderCandidate.sources, requested: [1, 0], complete: result.seriesOrderCandidate.complete };
  const number = { binding: result.valueBindings[0], value: 42 }, label = { closure: result.labelCandidates[1], value: 'changed category' };
  const staged = await prepareChartValues(source, [number], {}, DEFAULT_SESSION_LIMITS, async () => {}, 0, [label], [order]);
  assert.equal(staged.replacements.length, 2);
  const output = await replaceArchiveEntries(bytes, staged.replacements, { maxOutputBytes: 10_000_000 });
  const after = await inspectThroughImporter(output, true, DEFAULT_SESSION_LIMITS, undefined, false, true);
  assert.deepEqual(after.fields.series.map(item => [item.name, item.values]), [['Second & series', [4, 5, 6]], ['First _x005F_', [42, 2, 3]]]);
  assert.equal(after.fields.categories[1], 'changed category');
  assert.deepEqual(after.seriesOrderCandidate.sources.map(item => item.identity), [1, 0]);
  const chartXmlBytes = readGeneratedZip(output).get('ppt/charts/chart1.xml');
  const physicalIds = [...chartXmlBytes.toString().matchAll(/<c:ser><c:idx val="(\d+)"/gu)].map(match => Number(match[1]));
  assert.deepEqual(physicalIds, [1, 0]);
  await verifyChartSeriesOrderRebound(order.sources, order.requested, after.seriesOrderCandidate.sources, order.complete, after.seriesOrderCandidate.complete, 100, async () => {}, staged.movedElementIndexes.get(order.part));
  const beforeParts = readGeneratedZip(bytes), afterParts = readGeneratedZip(output);
  for (const [name, part] of beforeParts) if (!['ppt/charts/chart1.xml', 'ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx'].includes(name)) assert.deepEqual(afterParts.get(name), part, name);
  await assert.rejects(prepareChartValues(source, [], {}, DEFAULT_SESSION_LIMITS, async () => {}, 0, [], [order, order]), /duplicate/);
  await assert.rejects(prepareChartValues(source, [], {}, DEFAULT_SESSION_LIMITS, async () => {}, 0, [], [{ ...order, complete: false }]), { code: 'save-candidate-invalid' });
  await assert.rejects(prepareChartValues(source, [], {}, { ...DEFAULT_SESSION_LIMITS, maxSaveValidationWork: 1 }, async () => {}, 0, [], [order]), { code: 'save-limit-exceeded' });
  const bad = structuredClone(order); bad.sources[0].orderLexical = '9';
  await assert.rejects(prepareChartValues(source, [], {}, DEFAULT_SESSION_LIMITS, async () => {}, 0, [], [bad]), { code: 'save-candidate-invalid' });
});

test('research staging reorders a column chart and remaps proven legend indexes', async () => {
  const input = await readFile(new URL('../../../tests/fixtures/chart-order-legend-before.pptx', import.meta.url));
  const expected = await readFile(new URL('../../../tests/fixtures/chart-order-legend-remapped.pptx', import.meta.url));
  const inspected = await inspectThroughImporter(input, false, DEFAULT_SESSION_LIMITS, undefined, false, true);
  assert.equal(inspected.seriesOrderCandidate.complete, false);
  const order = { part: '/ppt/charts/chart1.xml', sources: inspected.seriesOrderCandidate.sources, requested: [1, 0], complete: true, researchColumnLegend: true };
  const source = await readOpcSource(input, {}, false, true), staged = await prepareChartValues(source, [], {}, DEFAULT_SESSION_LIMITS, async () => {}, 0, [], [order]);
  const actual = readGeneratedZip(await replaceArchiveEntries(input, staged.replacements, { maxOutputBytes: 10_000_000 }));
  const reference = readGeneratedZip(expected);
  for (const [name, bytes] of reference) assert.deepEqual(actual.get(name), bytes, name);
  const invalid = structuredClone(order); invalid.sources[0].identity = 123;
  await assert.rejects(prepareChartValues(source, [], {}, DEFAULT_SESSION_LIMITS, async () => {}, 0, [], [invalid]), { code: 'save-candidate-invalid' });
});

test('series-name candidates require direct ownership and exclusive workbook references', async () => {
  const source = await readFile(new URL('../../../tests/fixtures/chart-label-source.pptx', import.meta.url));
  const inspect = bytes => inspectThroughImporter(bytes, true, DEFAULT_SESSION_LIMITS, undefined, true);
  const result = await inspect(source);
  assert.deepEqual(result.seriesNameCandidates.map(item => [item.logicalKey, item.status, item.uses.length]), [['series:0', 'bound', 1], ['series:1', 'bound', 1]]);
  assert.equal(result.seriesNameCandidates[0].uses[0].projected, 'First _x005F_');
  const chart = readGeneratedZip(source).get('ppt/charts/chart1.xml').toString();
  const alias = chart.replace('Sheet1!$C$1', 'Sheet1!$B$1'); assert.notEqual(alias, chart);
  const changed = await replaceArchiveEntries(source, [{ name: 'ppt/charts/chart1.xml', bytes: new TextEncoder().encode(alias) }], { maxOutputBytes: 10_000_000 });
  assert.ok((await inspect(changed)).seriesNameCandidates.every(item => item.status === 'unmapped'));
  const broken = await inspectThroughImporter(source, true, DEFAULT_SESSION_LIMITS, 'tx', true);
  assert.ok(broken.seriesNameCandidates.every(item => item.status === 'unmapped'));
  assert.equal((await inspectThroughImporter(source)).seriesNameCandidates, undefined);
  for (const strict of [false, true]) {
    const categories = ['one', 'two'], data = series({ refs: true, name: 'Series', categories, values: [0, -2.5] });
    const input = chartBytes({ strict, xml: chartXml({ data, external: true }), book: workbook({ name: 'Series', categories, values: [0, -2.5] }) });
    const names = (await inspect(input)).seriesNameCandidates;
    assert.deepEqual(names.map(item => [item.logicalKey, item.status]), [['series:0', 'bound']]);
  }
  const unknown = chart.replace('<c:tx><c:strRef>', '<c:tx future="1"><c:strRef>'); assert.notEqual(unknown, chart);
  const unknownBytes = await replaceArchiveEntries(source, [{ name: 'ppt/charts/chart1.xml', bytes: new TextEncoder().encode(unknown) }], { maxOutputBytes: 10_000_000 });
  assert.equal((await inspect(unknownBytes)).seriesNameCandidates[0].status, 'unmapped');
});

test('scatter and bubble series-name candidates use direct workbook sources', async () => {
  for (const chartType of ['scatter', 'bubble']) {
    const bytes = await readFile(new URL(`../../../tests/fixtures/${chartType}-name-source.pptx`, import.meta.url));
    const result = await inspectThroughImporter(bytes, true, DEFAULT_SESSION_LIMITS, undefined, true);
    assert.deepEqual(result.seriesNameCandidates.map(item => [item.logicalKey, item.status]), [['series:0', 'bound']], chartType);
    assert.equal(result.seriesNameCandidates[0].uses[0].projected, 'Source & name');
    assert.equal(result.seriesNameCandidates[0].uses[0].cell.reference, 'B1');
    const broken = await inspectThroughImporter(bytes, true, DEFAULT_SESSION_LIMITS, 'tx', true);
    assert.equal(broken.seriesNameCandidates[0].status, 'unmapped');
  }
});

test('real chart import produces private label candidates without granting save authority', async () => {
  const source = await readFile(new URL('../../../tests/fixtures/chart-label-source.pptx', import.meta.url));
  const result = await inspectThroughImporter(source);
  assert.deepEqual(result.labelCandidates.map(item => [item.logicalKey, item.status, item.uses.length]), [['category:0', 'bound', 2], ['category:1', 'bound', 2], ['category:2', 'bound', 2]]);
  assert.equal(result.labelCandidates[1].uses[0].cache.lexical, 'literal _x005F_x0041_');
  assert.equal(result.labelCandidates[1].uses[0].cell.lexical, 'literal _x005F_x0041_');
  assert.equal((await inspectThroughImporter(source, false)).labelCandidates, undefined);
  await assert.rejects(inspectThroughImporter(source, true, { ...DEFAULT_SESSION_LIMITS, maxChartBindingWork: 1 }), error => error.code === 'import-limit-exceeded');
});

test('label candidates require the complete original series-to-chart ownership chain', async () => {
  const source = await readFile(new URL('../../../tests/fixtures/chart-label-source.pptx', import.meta.url));
  for (const ancestor of ['cat', 'ser', 'barChart', 'plotArea', 'chart']) {
    const result = await inspectThroughImporter(source, true, DEFAULT_SESSION_LIMITS, ancestor);
    assert.ok(result.labelCandidates.every(item => item.status === 'unmapped'), ancestor);
  }
});

test('private staging updates every category cache and its shared workbook cell atomically', async () => {
  const bytes = await readFile(new URL('../../../tests/fixtures/chart-label-source.pptx', import.meta.url));
  const result = await inspectThroughImporter(bytes), closure = result.labelCandidates[1];
  const source = await readOpcSource(bytes, {}, false, true), value = ' A & _x0041_ 😀\u0001 ';
  const numeric = result.valueBindings[0];
  assert.equal(numeric.status, 'bound');
  const staged = await prepareChartValues(source, [{ binding: numeric, value: 42.5 }], {}, DEFAULT_SESSION_LIMITS, async () => {}, 0, [{ closure, value }]);
  assert.equal(staged.replacements.length, 2);
  const output = await replaceArchiveEntries(bytes, staged.replacements, { maxOutputBytes: 10_000_000 });
  const after = await inspectThroughImporter(output);
  assert.equal(after.fields.categories[1], value);
  assert.equal(after.fields.series[0].values[0], 42.5);
  assert.equal(after.labelCandidates[1].status, 'bound');
  assert.equal(after.labelCandidates[1].uses.length, 2);
  assert.ok(after.labelCandidates[1].uses.every(use => use.projected === value));
  await assert.rejects(prepareChartValues(source, [], {}, DEFAULT_SESSION_LIMITS, async () => {}, 0, [{ closure, value }, { closure, value }]), /competing/);
  const invalid = structuredClone(closure); invalid.uses[1].cell.lexical = 'different';
  await assert.rejects(prepareChartValues(source, [], {}, DEFAULT_SESSION_LIMITS, async () => {}, 0, [{ closure: invalid, value }]), /conflict/);
  const noSpace = structuredClone(closure); for (const use of noSpace.uses) { use.cell.preserveSpace = false; delete use.cell.space; }
  await assert.rejects(prepareChartValues(source, [], {}, DEFAULT_SESSION_LIMITS, async () => {}, 0, [{ closure: noSpace, value }]), /presence|match/i);
});

test('private snapshot analysis binds only matching chart category identities and baselines', async () => {
  const bytes = await readFile(new URL('../../../tests/fixtures/chart-label-source.pptx', import.meta.url));
  const result = await inspectThroughImporter(bytes), closure = result.labelCandidates[1];
  const baseline = { slides: [{ id: 'slide', notes: [], elements: [{ id: 'chart', type: 'chart', categories: [...result.fields.categories], series: [{ values: [2] }] }] }] };
  const candidate = structuredClone(baseline); candidate.slides[0].elements[0].categories[1] = 'changed';
  const binding = { slideId: 'slide', elementId: 'chart', categoryIndex: 1, baselineValue: baseline.slides[0].elements[0].categories[1], closure };
  const analyze = (next, labels = [binding]) => analyzeTextEdits(baseline, next, [], [], DEFAULT_SESSION_LIMITS, async () => {}, [], undefined, [], [], [], labels);
  const accepted = await analyze(candidate);
  assert.equal(accepted.diagnostics.length, 0); assert.equal(accepted.chartLabels[0].value, 'changed');
  for (const bad of [{ ...binding, baselineValue: 'stale' }, { ...binding, elementId: 'other' }, { ...binding, closure: { ...closure, status: 'unmapped', reason: 'unsafe' } }]) {
    const rejected = await analyze(candidate, [bad]); assert.equal(rejected.chartLabels, undefined); assert.equal(rejected.diagnostics.length, 1);
  }
  const unrelated = structuredClone(candidate); unrelated.slides[0].elements[0].series[0].values[0] = 3;
  assert.equal((await analyze(unrelated)).diagnostics.length, 1);
  await assert.rejects(analyze(candidate, [binding, binding]), /Duplicate chart label/);
});

test('content projection associates private label closures with the actual chart identity', async () => {
  const bytes = await readFile(new URL('../../../tests/fixtures/chart-label-source.pptx', import.meta.url));
  const context = await structureContext(bytes, undefined, true), inspection = await readStructureFromContext(context);
  const checkpoint = async () => {}, projection = await projectStructure(inspection, DEFAULT_SESSION_LIMITS.maxReportFindings, checkpoint);
  const bindings = [], args = Array(35).fill(undefined);
  [args[0], args[1], args[2], args[3], args[4], args[5], args[34]] = [projection.document, projection.report, context, DEFAULT_SESSION_LIMITS, checkpoint, {}, bindings];
  await projectContent(...args);
  const chart = projection.document.slides[0].elements.find(element => element.type === 'chart');
  assert.ok(chart); assert.equal(bindings.length, 3);
  for (const [index, binding] of bindings.entries()) {
    assert.equal(binding.slideId, projection.document.slides[0].id); assert.equal(binding.elementId, chart.id);
    assert.equal(binding.categoryIndex, index); assert.equal(binding.baselineValue, chart.categories[index]); assert.equal(binding.closure.status, 'bound');
  }
});

test('private candidate pipeline stages and rebinds category labels against the full Core snapshot', async () => {
  const bytes = await readFile(new URL('../../../tests/fixtures/chart-label-source.pptx', import.meta.url));
  const context = await structureContext(bytes, undefined, true), inspection = await readStructureFromContext(context);
  const checkpoint = async () => {}, projection = await projectStructure(inspection, DEFAULT_SESSION_LIMITS.maxReportFindings, checkpoint);
  const bindings = [], args = Array(35).fill(undefined);
  [args[0], args[1], args[2], args[3], args[4], args[5], args[34]] = [projection.document, projection.report, context, DEFAULT_SESSION_LIMITS, checkpoint, {}, bindings];
  await projectContent(...args);
  const baseline = snapshot(projection.document, DEFAULT_SESSION_LIMITS.maxSnapshotValues), candidate = structuredClone(baseline);
  const chart = candidate.slides[0].elements.find(element => element.type === 'chart'); chart.categories[1] = 'new & _x0041_';
  const analysis = await analyzeTextEdits(baseline, candidate, [], [], DEFAULT_SESSION_LIMITS, checkpoint, [], undefined, [], [], [], bindings);
  assert.equal(analysis.diagnostics.length, 0);
  const prepared = await prepareTextCandidate(bytes, candidate, analysis, {}, DEFAULT_SESSION_LIMITS, checkpoint);
  assert.ok(prepared.bytes); const result = await inspectThroughImporter(prepared.bytes);
  assert.equal(result.fields.categories[1], chart.categories[1]); assert.equal(result.labelCandidates[1].status, 'bound');
  const wrong = structuredClone(analysis); wrong.chartLabels[0].binding.closure.uses[0].cache.elementIndex++;
  await assert.rejects(prepareTextCandidate(bytes, candidate, wrong, {}, DEFAULT_SESSION_LIMITS, checkpoint), /identity|expected|source|match/i);
});

test('private full candidate rebinds reordered series with mixed name, value and category edits', async () => {
  const bytes = await readFile(new URL('../../../tests/fixtures/chart-label-source.pptx', import.meta.url));
  const context = await structureContext(bytes, undefined, true), inspection = await readStructureFromContext(context), checkpoint = async () => {};
  const projection = await projectStructure(inspection, DEFAULT_SESSION_LIMITS.maxReportFindings, checkpoint), labels = [], numbers = [], orders = [], args = Array(37).fill(undefined);
  [args[0], args[1], args[2], args[3], args[4], args[5], args[9], args[34], args[35], args[36]] = [projection.document, projection.report, context, DEFAULT_SESSION_LIMITS, checkpoint, {}, numbers, labels, true, orders];
  await projectContent(...args); assert.equal(orders.length, 1); assert.equal(orders[0].complete, true);
  const baseline = snapshot(projection.document, DEFAULT_SESSION_LIMITS.maxSnapshotValues), scalarCandidate = structuredClone(baseline);
  const chart = scalarCandidate.slides[0].elements.find(element => element.type === 'chart'); chart.series[0].name = 'renamed'; chart.series[0].values[0] = 42; chart.categories[1] = 'category';
  const candidate = structuredClone(scalarCandidate); candidate.slides[0].elements.find(element => element.type === 'chart').series.reverse();
  const request = { slideId: orders[0].slideId, elementId: orders[0].elementId, identities: [1, 0] };
  const mapped = await analyzeChartOrders(baseline, candidate, orders, [request], DEFAULT_SESSION_LIMITS.maxSaveAnalysisWork, checkpoint);
  assert.ok(Number.isSafeInteger(mapped.work) && mapped.work > 0);
  await assert.rejects(analyzeChartOrders(baseline, candidate, orders, [request], mapped.work - 1, checkpoint), { code: 'save-limit-exceeded' });
  assert.equal((await analyzeChartOrders(baseline, candidate, orders, [request], mapped.work, checkpoint)).work, mapped.work);
  const analysis = await analyzeTextEdits(baseline, mapped.candidateForScalars, [], [], DEFAULT_SESSION_LIMITS, checkpoint, numbers, undefined, [], [], [], labels);
  assert.deepEqual(analysis.diagnostics, []); assert.equal(analysis.chartValues.length, 1); assert.equal(analysis.chartLabels.length, 2);
  assert.deepEqual(candidate.slides[0].elements.find(element => element.type === 'chart').series.map(item => item.name), ['Second & series', 'renamed']);
  const mixed = { ...analysis, chartOrders: mapped.proposals, changes: [...analysis.changes, ...mapped.changes] };
  const prepared = await prepareTextCandidate(bytes, candidate, mixed, {}, DEFAULT_SESSION_LIMITS, checkpoint);
  assert.ok(prepared.bytes); const after = await inspectThroughImporter(prepared.bytes, true, DEFAULT_SESSION_LIMITS, undefined, true, true);
  assert.deepEqual(after.fields.series.map(item => [item.name, item.values]), [['Second & series', [4, 5, 6]], ['renamed', [42, 2, 3]]]);
  const wrong = structuredClone(mixed); wrong.chartOrders[0].requested = [0, 1];
  await assert.rejects(prepareTextCandidate(bytes, candidate, wrong, {}, DEFAULT_SESSION_LIMITS, checkpoint), /snapshot|differs/i);
  for (const invalid of [{ ...request, identities: [0, 0] }, { ...request, identities: [0, 1] }, { ...request, elementId: 'missing' }]) {
    await assert.rejects(analyzeChartOrders(baseline, candidate, orders, [invalid], DEFAULT_SESSION_LIMITS.maxSaveAnalysisWork, checkpoint), { code: 'save-candidate-invalid' });
  }
  await assert.rejects(analyzeChartOrders(baseline, candidate, orders, [request, request], DEFAULT_SESSION_LIMITS.maxSaveAnalysisWork, checkpoint), { code: 'save-candidate-invalid' });
  await assert.rejects(analyzeChartOrders(baseline, candidate, orders, [request], 1, checkpoint), { code: 'save-limit-exceeded' });
  const missing = structuredClone(candidate); missing.slides[0].elements.find(element => element.type === 'chart').series.pop();
  await assert.rejects(analyzeChartOrders(baseline, missing, orders, [request], DEFAULT_SESSION_LIMITS.maxSaveAnalysisWork, checkpoint), { code: 'save-candidate-invalid' });
  const cancelled = new Error('cancelled'); await assert.rejects(analyzeChartOrders(baseline, candidate, orders, [request], DEFAULT_SESSION_LIMITS.maxSaveAnalysisWork, async () => { throw cancelled; }), error => error === cancelled);
});

test('research full candidate rebinds legend identity and mixed series edits across three generations', async () => {
  let bytes = await readFile(new URL('../../../tests/fixtures/chart-order-legend-before.pptx', import.meta.url));
  for (let round = 1; round <= 3; round++) {
    const context = await structureContext(bytes, undefined, true), inspection = await readStructureFromContext(context), checkpoint = async () => {};
    const projection = await projectStructure(inspection, DEFAULT_SESSION_LIMITS.maxReportFindings, checkpoint), labels = [], numbers = [], orders = [], args = Array(38).fill(undefined);
    [args[0], args[1], args[2], args[3], args[4], args[5], args[9], args[34], args[35], args[36], args[37]] = [projection.document, projection.report, context, DEFAULT_SESSION_LIMITS, checkpoint, {}, numbers, labels, true, orders, true];
    await projectContent(...args); assert.equal(orders[0].complete, true); assert.equal(orders[0].researchColumnLegend, true);
    const baseline = snapshot(projection.document, DEFAULT_SESSION_LIMITS.maxSnapshotValues), candidate = structuredClone(baseline);
    const chart = candidate.slides[0].elements.find(item => item.type === 'chart');
    chart.series[0].name = `legend round ${round}`; chart.series[0].values[0] = 20 + round; chart.categories[1] = `category ${round}`; chart.series.reverse();
    const binding = orders[0], identities = [...binding.sources].sort((a, b) => a.order - b.order).map(item => item.identity).reverse();
    const mapped = await analyzeChartOrders(baseline, candidate, orders, [{ slideId: binding.slideId, elementId: binding.elementId, identities }], DEFAULT_SESSION_LIMITS.maxSaveAnalysisWork, checkpoint);
    const scalar = await analyzeTextEdits(baseline, mapped.candidateForScalars, [], [], DEFAULT_SESSION_LIMITS, checkpoint, numbers, undefined, [], [], [], labels);
    assert.deepEqual(scalar.diagnostics, []);
    const output = await prepareTextCandidate(bytes, candidate, { ...scalar, chartOrders: mapped.proposals, changes: [...scalar.changes, ...mapped.changes] }, {}, DEFAULT_SESSION_LIMITS, checkpoint);
    assert.ok(output.bytes); bytes = output.bytes;
    const raw = readGeneratedZip(bytes).get('ppt/charts/chart1.xml').toString();
    assert.ok(raw.includes(`<c:legendEntry><c:idx val="${round % 2}"/><c:delete val="1"/></c:legendEntry>`));
  }
});

test('research legend binding does not override other unsafe source semantics', async () => {
  const bytes = await readFile(new URL('../../../tests/fixtures/chart-order-legend-before.pptx', import.meta.url)), parts = readGeneratedZip(bytes);
  const name = 'ppt/charts/chart1.xml', xml = parts.get(name).toString();
  for (const changed of [xml.replace('</c:chartSpace>', '<c:extLst/></c:chartSpace>'), xml.replace('<c:idx val="0"/><c:delete val="1"/>', '<c:idx val="0"/><c:delete val="1"/><c:txPr/>')]) {
    assert.notEqual(changed, xml);
    const input = await replaceArchiveEntries(bytes, [{ name, bytes: new TextEncoder().encode(changed) }], { maxOutputBytes: 10_000_000 });
    const context = await structureContext(input, undefined, true), inspection = await readStructureFromContext(context), checkpoint = async () => {};
    const projection = await projectStructure(inspection, DEFAULT_SESSION_LIMITS.maxReportFindings, checkpoint), orders = [], args = Array(38).fill(undefined);
    [args[0], args[1], args[2], args[3], args[4], args[5], args[36], args[37]] = [projection.document, projection.report, context, DEFAULT_SESSION_LIMITS, checkpoint, {}, orders, true];
    await projectContent(...args); assert.equal(orders[0].complete, false); assert.ok(orders[0].localDependencyReason);
  }
});

test('PowerPoint resaved chart extensions keep research legend ordering unmapped', async () => {
  const bytes = await readFile(new URL('../../../tests/fixtures/chart-order-legend-powerpoint-resave.pptx', import.meta.url));
  const context = await structureContext(bytes, undefined, true), inspection = await readStructureFromContext(context), checkpoint = async () => {};
  const projection = await projectStructure(inspection, DEFAULT_SESSION_LIMITS.maxReportFindings, checkpoint), orders = [], args = Array(38).fill(undefined);
  [args[0], args[1], args[2], args[3], args[4], args[5], args[36], args[37]] = [projection.document, projection.report, context, DEFAULT_SESSION_LIMITS, checkpoint, {}, orders, true];
  await projectContent(...args); assert.equal(orders.length, 1); assert.equal(orders[0].complete, false); assert.match(orders[0].localDependencyReason, /extensions|foreign/);
});

test('chart-order graph inspection rejects shared and incomplete relationship sources', async () => {
  const bytes = await readFile(new URL('../../../tests/fixtures/chart-label-source.pptx', import.meta.url));
  const binding = { slideId: 'slide', elementId: 'chart', part: '/ppt/charts/chart1.xml', slidePart: '/ppt/slides/slide1.xml', sources: [], complete: true };
  const graph = await inspectClosure(bytes, { roots: [binding.part] }), checkpoint = async () => {};
  await verifyChartOrderRelationshipClosure(binding, graph, checkpoint);
  const wrongSlide = { ...binding, slidePart: '/ppt/slides/other.xml' };
  await assert.rejects(verifyChartOrderRelationshipClosure(wrongSlide, graph, checkpoint), /owning slide/);
  const incomplete = structuredClone(graph); incomplete.coverage.incomingComplete = false;
  await assert.rejects(verifyChartOrderRelationshipClosure(binding, incomplete, checkpoint), /incomplete/);
  const parts = readGeneratedZip(bytes), relName = 'ppt/slides/_rels/slide1.xml.rels', rels = parts.get(relName).toString();
  const altered = rels.replace('</Relationships>', '<Relationship Id="rIdExtra" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>'); assert.notEqual(altered, rels);
  const shared = await replaceArchiveEntries(bytes, [{ name: relName, bytes: new TextEncoder().encode(altered) }], { maxOutputBytes: 10_000_000 });
  const sharedGraph = await inspectClosure(shared, { roots: [binding.part] });
  await assert.rejects(verifyChartOrderRelationshipClosure(binding, sharedGraph, checkpoint), /exactly one incoming/);
  const wrongOutgoing = structuredClone(graph); wrongOutgoing.relationships[0].type = 'urn:unknown';
  await assert.rejects(verifyChartOrderRelationshipClosure(binding, wrongOutgoing, checkpoint), /unsupported relationship/);
  const cancelled = new Error('cancelled'); await assert.rejects(verifyChartOrderRelationshipClosure(binding, graph, async () => { throw cancelled; }), error => error === cancelled);
});

test('private candidate inspection follows both OOXML namespace dialects', async () => {
  for (const strict of [false, true]) {
    const categories = ['one', 'two'], data = series({ refs: true, name: 'Series', categories, values: [0, -2.5] });
    const input = chartBytes({ strict, xml: chartXml({ data, external: true }), book: workbook({ name: 'Series', categories, values: [0, -2.5] }) });
    const result = await inspectThroughImporter(input);
    assert.deepEqual(result.labelCandidates.map(item => [item.logicalKey, item.status, item.uses.length]), [['category:0', 'bound', 1], ['category:1', 'bound', 1]]);
  }
});

test('real source inspection refuses candidates when chart references or workbook strings are unsafe', async () => {
  const source = await readFile(new URL('../../../tests/fixtures/chart-label-source.pptx', import.meta.url)), parts = readGeneratedZip(source);
  const chart = parts.get('ppt/charts/chart1.xml').toString();
  const extraFormula = chart.replace('</c:chartSpace>', '<c:extLst><c:ext uri="future"><c:f>Sheet1!$A$3</c:f></c:ext></c:extLst></c:chartSpace>');
  const unknown = await replaceArchiveEntries(source, [{ name: 'ppt/charts/chart1.xml', bytes: new TextEncoder().encode(extraFormula) }], { maxOutputBytes: 10_000_000 });
  const result = await inspectThroughImporter(unknown); assert.ok(result.labelCandidates.every(item => item.status === 'unmapped'));
  const nameAlias = chart.replace('Sheet1!$B$1', 'Sheet1!$A$2');
  const alias = await replaceArchiveEntries(source, [{ name: 'ppt/charts/chart1.xml', bytes: new TextEncoder().encode(nameAlias) }], { maxOutputBytes: 10_000_000 });
  const aliased = await inspectThroughImporter(alias); assert.match(aliased.labelCandidates.find(item => item.logicalKey === 'category:0').reason, /unaccounted/);
  const bookName = 'ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx', book = parts.get(bookName), sheet = readGeneratedZip(book).get('xl/worksheets/sheet1.xml').toString();
  const old = '<is><t xml:space="preserve">literal _x005F_x0041_</t></is>', rich = '<is><r><t>literal _x005F_x0041_</t></r></is>';
  assert.ok(sheet.includes(old));
  const changedBook = await replaceArchiveEntries(book, [{ name: 'xl/worksheets/sheet1.xml', bytes: new TextEncoder().encode(sheet.replace(old, rich)) }], { maxOutputBytes: 10_000_000 });
  const changed = await replaceArchiveEntries(source, [{ name: bookName, bytes: changedBook }], { maxOutputBytes: 10_000_000 });
  const richResult = await inspectThroughImporter(changed); assert.equal(richResult.labelCandidates.find(item => item.logicalKey === 'category:1').status, 'unmapped');
  const conflictBook = await replaceArchiveEntries(book, [{ name: 'xl/worksheets/sheet1.xml', bytes: new TextEncoder().encode(sheet.replace('literal _x005F_x0041_', 'other')) }], { maxOutputBytes: 10_000_000 });
  const conflict = await replaceArchiveEntries(source, [{ name: bookName, bytes: conflictBook }], { maxOutputBytes: 10_000_000 });
  const mismatch = await inspectThroughImporter(conflict); assert.match(mismatch.labelCandidates.find(item => item.logicalKey === 'category:1').reason, /disagree/);
});

test('candidate source ancestry rejects unknown wrappers, fields and value metadata', async () => {
  const source = await readFile(new URL('../../../tests/fixtures/chart-label-source.pptx', import.meta.url));
  const chart = readGeneratedZip(source).get('ppt/charts/chart1.xml').toString();
  const cases = [
    ['<c:cat><c:strRef>', '<c:cat future="1"><c:strRef>'],
    ['<c:cat><c:strRef>', '<c:cat><c:strRef future="1">'],
    ['<c:f>Sheet1!$A$2:$A$4</c:f>', '<c:f future="1">Sheet1!$A$2:$A$4</c:f>'],
    ['<c:strCache><c:ptCount val="3"/>', '<c:strCache future="1"><c:ptCount val="3"/>'],
    ['<c:strCache><c:ptCount val="3"/>', '<c:strCache><c:ptCount val="3" future="1"/>'],
    ['<c:pt idx="0"><c:v> Q1 </c:v>', '<c:pt idx="0" future="1"><c:v> Q1 </c:v>'],
    ['<c:pt idx="0"><c:v> Q1 </c:v>', '<c:pt idx="0"><c:v future="1"> Q1 </c:v>'],
    ['<c:strCache><c:ptCount val="3"/>', '<c:strCache><c:extLst/><c:ptCount val="3"/>'],
  ];
  for (const [old, replacement] of cases) {
    assert.ok(chart.includes(old));
    const changed = await replaceArchiveEntries(source, [{ name: 'ppt/charts/chart1.xml', bytes: new TextEncoder().encode(chart.replace(old, replacement)) }], { maxOutputBytes: 10_000_000 });
    const result = await inspectThroughImporter(changed);
    assert.equal(result.labelCandidates.find(item => item.logicalKey === 'category:0').status, 'unmapped', replacement);
  }
});
