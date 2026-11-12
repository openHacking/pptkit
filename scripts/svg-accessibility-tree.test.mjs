import test from 'node:test';
import assert from 'node:assert/strict';
import { chromium } from '@playwright/test';
import { createPresentation, normalizePresentation, validatePresentation } from '../packages/core/dist/index.js';
import { renderIncrementalSvg, renderLayoutToSvg, renderPresentationToSvg } from '../packages/svg-renderer/dist/index.js';
import { resolveLayout } from '../packages/layout/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';
import { readGeneratedZip } from '../tests/support/read-generated-zip.mjs';

test('accessibility batch edits and inverse change the observed SVG tree and restore reading order', async t => {
  const deck = createPresentation({ id: 'deck', size: { width: 400, height: 300 } });
  const box = { x: 0, y: 0, width: 120, height: 80 };
  const slide = deck.addSlide({ id: 'page', elements: [
    { id: 'group', type: 'group', box, coordinateSize: { width: 120, height: 80 }, accessibility: { description: 'Group before' }, children: [
      { id: 'child', type: 'shape', shape: 'rect', box, accessibility: { description: 'Child before' } },
    ] },
    { id: 'other', type: 'shape', shape: 'rect', box, accessibility: { description: 'Other' } },
  ] });
  slide.setReadingOrder([{ elementIds: ['other', 'group'] }, { groupId: 'group', elementIds: ['child'] }]);
  const browser = await chromium.launch({ headless: true }); t.after(() => browser.close());
  const page = await browser.newPage(), cdp = await page.context().newCDPSession(page);
  async function names() {
    await page.setContent((await renderPresentationToSvg(deck)).slides[0].svg);
    const tree = await cdp.send('Accessibility.getFullAXTree');
    return tree.nodes.filter(node => !node.ignored).map(node => node.name?.value);
  }
  assert.ok((await names()).includes('Child before'));
  const undo = deck.applyEditOperation({ schemaVersion: 'pptkit-edit-v1', documentId: 'deck', changes: [
    { type: 'element-accessibility', slideId: 'page', elementId: 'group', expected: { description: 'Group before', decorative: null }, value: { description: 'Group after', decorative: true } },
    { type: 'element-accessibility', slideId: 'page', elementId: 'child', expected: { description: 'Child before', decorative: null }, value: { description: 'Child after', decorative: null } },
  ] }).inverse;
  const hiddenNames = await names(); assert.ok(!hiddenNames.includes('Child after')); assert.ok(!hiddenNames.includes('Group after'));
  assert.deepEqual(slide.readingOrder, [{ elementIds: ['other'] }]);
  deck.applyEditOperation(undo); const restored = await names();
  assert.ok(restored.includes('Child before')); assert.ok(restored.includes('Group before'));
  assert.deepEqual(slide.readingOrder, [{ elementIds: ['other', 'group'] }, { groupId: 'group', elementIds: ['child'] }]);
});

test('Chromium accessibility tree exposes SVG object descriptions and hides decorative subtrees', async t => {
  const deck = createPresentation({ size: { width: 800, height: 600 } }), box = { x: 0, y: 0, width: 120, height: 80 };
  deck.registerAsset({ id: 'asset', kind: 'image', source: { type: 'bytes', encoding: 'base64', value: 'iVBORw0KGgoAAAANSUhEUgAAAAgAAAAECAIAAAA8r+mnAAAAIUlEQVR4nGNQCb8GRze90uCIAaeE0tE4OPp1RhSOcEoAAP18L9GcKmUbAAAAAElFTkSuQmCC' }, mimeType: 'image/png', width: 8, height: 4 });
  const group = (id, decorative, childName) => ({ id, type: 'group', box, coordinateSize: { width: 120, height: 80 }, accessibility: { decorative, description: `${id} description` }, children: [{ id: `${id}-child`, type: 'shape', shape: 'rect', box, accessibility: { description: childName } }] });
  deck.addSlide({ id: 'page', elements: [
    { id: 'image', type: 'image', box, assetId: 'asset', accessibility: { description: 'Image description' } },
    { id: 'chart', type: 'chart', box, chartType: 'bar', categories: ['A'], series: [{ name: 'S', values: [1] }], accessibility: { description: 'Chart description' } },
    { id: 'table', type: 'table', box, columns: [120], rows: [{ cells: [{ content: 'Cell content' }] }], accessibility: { description: 'Table description' } },
    group('ordinary-group', false, 'Ordinary child'), group('decorative-group', true, 'Suppressed child'),
    { ...group('decorative-action', true, 'Suppressed action child'), action: { type: 'url', url: 'https://example.com', tooltip: 'Decorative link action' }, children: [{ id: 'action-child', type: 'shape', shape: 'rect', box, action: { type: 'url', url: 'https://example.com', tooltip: 'Hidden child action' } }] },
    { id: 'hidden-image', type: 'image', box, assetId: 'asset', accessibility: { decorative: true, description: 'Suppressed image' } },
  ] });
  const svg = (await renderPresentationToSvg(deck)).slides[0].svg;
  const browser = await chromium.launch({ headless: true }); t.after(() => browser.close());
  const page = await browser.newPage(), requests = []; await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
  await page.setContent(svg);
  const cdp = await page.context().newCDPSession(page), tree = await cdp.send('Accessibility.getFullAXTree');
  const nodes = tree.nodes.filter(node => !node.ignored).map(node => ({ role: node.role?.value, name: node.name?.value }));
  for (const name of ['Image description', 'Chart description', 'Ordinary child']) assert.ok(nodes.some(node => node.role === 'image' && node.name === name), name);
  assert.equal(nodes.filter(node => node.name === 'ordinary-group description').length, 1);
  assert.ok(nodes.some(node => node.role === 'group' && node.name === 'ordinary-group description'));
  for (const name of ['decorative-group description', 'Suppressed child', 'Suppressed image']) assert.ok(!nodes.some(node => node.name === name), name);
  assert.ok(nodes.some(node => node.role === 'link' && node.name === 'Decorative link action'));
  assert.ok(!nodes.some(node => node.name === 'Hidden child action'));
  assert.equal(nodes.filter(node => node.role === 'table' && node.name === 'Table description').length, 1);
  assert.ok(!nodes.some(node => node.role === 'image' && node.name === 'Table description'));
  assert.ok(nodes.some(node => node.role === 'cell' && node.name === 'Cell content'));
  assert.deepEqual(requests, []);
});

test('Chromium accessibility-tree object order follows emitted nested SVG after move, copy and delete', async t => {
  const deck = createPresentation(), box = { x: 0, y: 0, width: 100, height: 50 };
  const slide = deck.addSlide({ id: 'page', elements: [
    { id: 'first', name: 'First', type: 'shape', shape: 'rect', box, accessibility: { description: 'First' } },
    { id: 'group', name: 'Group', type: 'group', box, coordinateSize: { width: 100, height: 50 }, accessibility: { description: 'Group' }, children: [{ id: 'child', name: 'Child', type: 'shape', shape: 'rect', box, accessibility: { description: 'Child' } }, { id: 'child2', name: 'Child 2', type: 'shape', shape: 'rect', box, accessibility: { description: 'Child 2' } }] },
    { id: 'last', name: 'Last', type: 'shape', shape: 'rect', box, accessibility: { description: 'Last' } },
  ] });
  const browser = await chromium.launch({ headless: true }); t.after(() => browser.close());
  const page = await browser.newPage(), cdp = await page.context().newCDPSession(page);
  async function observed() {
    await page.setContent((await renderPresentationToSvg(deck)).slides[0].svg);
    const tree = await cdp.send('Accessibility.getFullAXTree'), byId = new Map(tree.nodes.map(node => [node.nodeId, node]));
    const order = [], walk = id => { const node = byId.get(id); if (!node) return; if (!node.ignored && ['image', 'group'].includes(node.role?.value) && ['First', 'Group', 'Child', 'Child 2', 'Last', 'Copy group', 'Copy child', 'Copy child 2'].includes(node.name?.value)) order.push(node.name.value); for (const child of node.childIds ?? []) walk(child); };
    walk(tree.nodes[0].nodeId);
    const xml = readGeneratedZip((await generatePptx(deck)).bytes).get('ppt/slides/slide1.xml').toString();
    const pptxNames = [...xml.matchAll(/<p:cNvPr\b[^>]*\bname="([^"]*)"/g)].map(match => match[1]).filter(name => ['First', 'Group', 'Child', 'Child 2', 'Last', 'Copy group', 'Copy child', 'Copy child 2'].includes(name));
    assert.deepEqual(pptxNames, order);
    return order;
  }
  assert.deepEqual(await observed(), ['First', 'Group', 'Child', 'Child 2', 'Last']);
  slide.moveElement('last', 0);
  assert.deepEqual(await observed(), ['Last', 'First', 'Group', 'Child', 'Child 2']);
  slide.reorderElements(['group', 'last', 'first']);
  assert.deepEqual(await observed(), ['Group', 'Child', 'Child 2', 'Last', 'First']);
  slide.reorderGroupElements('group', ['child2', 'child']);
  assert.deepEqual(await observed(), ['Group', 'Child 2', 'Child', 'Last', 'First']);
  const duplicate = slide.duplicateElement('group'), changed = structuredClone(duplicate);
  changed.name = 'Copy group'; changed.accessibility.description = 'Copy group'; changed.children[0].name = 'Copy child 2'; changed.children[0].accessibility.description = 'Copy child 2'; changed.children[1].name = 'Copy child'; changed.children[1].accessibility.description = 'Copy child';
  slide.replaceElement(duplicate.id, changed); slide.removeElement('first');
  assert.notEqual(duplicate.id, 'group'); assert.notEqual(changed.children[0].id, 'child2'); assert.notEqual(changed.children[1].id, 'child');
  assert.deepEqual(await observed(), ['Group', 'Child 2', 'Child', 'Copy group', 'Copy child 2', 'Copy child', 'Last']);
});

test('explicit table headers reach Chromium through merged and legacy layout cells without changing PPTX formatting', async t => {
  const element = {
    id: 'measurements', type: 'table', name: 'Fallback table name',
    accessibility: { description: 'Revenue "by quarter" & metric' },
    box: { x: 20, y: 20, width: 240, height: 120 },
    columns: [60, 90, 90], rows: [
      { cells: [{ content: 'Metric', headerScope: 'column' }, { content: 'Quarter', headerScope: 'column', colSpan: 2 }] },
      { cells: [{ content: 'Revenue', headerScope: 'row', rowSpan: 2 }, { content: '20' }, { content: '30' }] },
      { cells: [{ content: '21' }, { content: '31' }] },
    ],
    style: { firstRow: { fill: { type: 'solid', color: '#abcdef' } } },
  };
  const make = input => {
    const deck = createPresentation();
    deck.addSlide({ id: 'metrics', elements: [input] });
    return deck;
  };
  const deck = make(element);
  assert.deepEqual(validatePresentation(deck), []);
  const normalized = normalizePresentation(deck).slides[0].elements[0];
  assert.deepEqual(normalized.rows.map(row => row.cells.map(cell => cell.headerScope)), [['column', 'column'], ['row', undefined, undefined], [undefined, undefined]]);
  const malformed = structuredClone(element);
  malformed.rows[0].cells[0].headerScope = 'col';
  assert.ok(validatePresentation(make(malformed)).some(d => d.code === 'invalid-table-header-scope' && d.path.endsWith('.rows.0.cells.0.headerScope')));
  const unmarked = structuredClone(element);
  for (const row of unmarked.rows) for (const cell of row.cells) delete cell.headerScope;

  const native = await generatePptx(deck);
  assert.deepEqual(native.bytes, (await generatePptx(make(unmarked))).bytes);
  assert.deepEqual(native.warnings.filter(w => w.code === 'table-header-semantics-not-exported').map(w => [w.slideId, w.elementId]), [['metrics', 'measurements']]);

  const browser = await chromium.launch();
  t.after(() => browser.close());
  const page = await browser.newPage();
  const cdp = await page.context().newCDPSession(page);
  async function roles(svg, expectedTableName = element.accessibility.description, standalone = false) {
    if (standalone) await page.goto(`data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`);
    else await page.setContent(svg);
    const tree = await cdp.send('Accessibility.getFullAXTree');
    const indexed = new Map(tree.nodes.map(node => [node.nodeId, node]));
    const tables = tree.nodes.filter(node => !node.ignored && node.role?.value === 'table');
    assert.deepEqual(tables.map(node => node.name?.value), expectedTableName === null ? [] : [expectedTableName]);
    assert.ok(!tree.nodes.some(node => !node.ignored && node.role?.value === 'image' && node.name?.value === element.accessibility.description));
    const belongsToTable = node => {
      for (let parent = indexed.get(node.parentId); parent; parent = indexed.get(parent.parentId)) if (parent === tables[0]) return true;
      return false;
    };
    const names = new Set(['Metric', 'Quarter', 'Revenue', '20', '30', '21', '31']);
    const cells = tree.nodes.filter(node => !node.ignored && names.has(node.name?.value) &&
      ['columnheader', 'rowheader', 'cell'].includes(node.role?.value));
    assert.equal(new Set(cells.map(node => node.name.value)).size, cells.length);
    for (const cell of cells) assert.ok(belongsToTable(cell), cell.name.value);
    return Object.fromEntries(cells.map(node => [node.name.value, node.role.value]));
  }
  const svg = (await renderPresentationToSvg(deck)).slides[0].svg;
  assert.match(svg, /data-pptkit-element-id="measurements"/);
  assert.match(svg, /role="columnheader" aria-colindex="2" aria-rowindex="1" aria-colspan="2"/);
  assert.match(svg, /role="rowheader" aria-colindex="1" aria-rowindex="2" aria-colspan="1" aria-rowspan="2"/);
  assert.match(svg, /role="cell" aria-colindex="2" aria-rowindex="3"/);
  const expected = { Metric: 'columnheader', Quarter: 'columnheader', Revenue: 'rowheader', 20: 'cell', 30: 'cell', 21: 'cell', 31: 'cell' };
  assert.deepEqual(await roles(svg), expected);
  const control = await roles((await renderPresentationToSvg(make(unmarked))).slides[0].svg);
  assert.deepEqual(control, Object.fromEntries(Object.keys(expected).map(name => [name, 'cell'])));
  const decorative = make({ ...element, accessibility: { decorative: true } });
  assert.deepEqual(await roles((await renderPresentationToSvg(decorative)).slides[0].svg, null), {});
  const named = make({ ...element, accessibility: undefined, action: { type: 'url', url: 'https://example.com', tooltip: 'Open report' } });
  assert.deepEqual(await roles((await renderPresentationToSvg(named)).slides[0].svg, element.name), expected);
  const linked = await cdp.send('Accessibility.getFullAXTree');
  assert.ok(linked.nodes.some(node => !node.ignored && node.role?.value === 'link' && node.name?.value === 'Open report'));
  const legacy = structuredClone(resolveLayout(deck)), legacyTable = legacy.slides[0].elements[0];
  delete legacyTable.cellLayout;
  delete legacyTable.borderGrid;
  const legacySvg = (await renderLayoutToSvg(legacy)).slides[0].svg;
  assert.match(legacySvg, /<th scope="col" rowspan="1" colspan="2"/);
  assert.match(legacySvg, /<th scope="row" rowspan="2" colspan="1"/);
  assert.deepEqual(await roles(legacySvg), expected);
  assert.deepEqual(await roles(svg, element.accessibility.description, true), expected);
});

test('paginated repeated and merged headers retain accessible roles on every slide', async t => {
  const deck = createPresentation({ size: { width: 300, height: 200 } });
  deck.addSlide({ id: 'source', elements: [{
    id: 'table', type: 'table', name: 'Report', box: { x: 0, y: 0, width: 100, height: 100 }, columns: [50, 50], rows: [
      { height: 10, cells: [{ content: 'Metric', headerScope: 'column' }, { content: 'Value', headerScope: 'column' }] },
      { height: 20, cells: [{ content: 'First', headerScope: 'row' }, { content: '1' }] },
      { height: 20, cells: [{ content: 'Second', headerScope: 'row' }, { content: '2' }] },
      { height: 20, cells: [{ content: 'Merged', headerScope: 'row', rowSpan: 2 }, { content: '3' }] },
      { height: 20, cells: [{ content: '4' }] },
    ],
  }] });
  const layout = { tablePagination: [{ elementId: 'table', area: { x: 0, y: 0, width: 100, height: 50 }, headerRows: 1 }] };
  const [preview, native] = await Promise.all([renderPresentationToSvg(deck, { layout }), generatePptx(deck, { layout })]);
  assert.equal(preview.slides.length, 2);
  assert.equal(native.warnings.filter(w => w.code === 'table-header-semantics-not-exported').length, 2);
  const browser = await chromium.launch({ headless: true });
  t.after(() => browser.close());
  const page = await browser.newPage(), cdp = await page.context().newCDPSession(page);
  const expected = [
    { Metric: 'columnheader', Value: 'columnheader', First: 'rowheader', Second: 'rowheader', 1: 'cell', 2: 'cell' },
    { Metric: 'columnheader', Value: 'columnheader', Merged: 'rowheader', 3: 'cell', 4: 'cell' },
  ];
  for (let i = 0; i < preview.slides.length; i++) {
    const svg = preview.slides[i].svg;
    await page.setContent(svg);
    const tree = await cdp.send('Accessibility.getFullAXTree');
    const cells = tree.nodes.filter(node => !node.ignored && ['columnheader', 'rowheader', 'cell'].includes(node.role?.value));
    assert.deepEqual(Object.fromEntries(cells.map(node => [node.name.value, node.role.value])), expected[i]);
    if (i === 1) assert.match(svg, /role="rowheader" aria-colindex="1" aria-rowindex="2" aria-colspan="1" aria-rowspan="2"/);
  }
});

test('nested SVG groups keep interactive and structured children under a single named group', async t => {
  const deck = createPresentation(), box = { x: 10, y: 20, width: 160, height: 90 };
  const table = { type: 'table', box, columns: [160], rows: [{ cells: [{ content: 'Quantity', headerScope: 'column' }] }], name: 'Orders table' };
  const chart = { type: 'chart', chartType: 'bar', box, categories: ['East'], series: [{ name: 'Orders', values: [2] }], accessibility: { description: 'Orders chart' } };
  const inner = { type: 'group', name: 'Child group', box, coordinateSize: box, children: [table, chart] };
  deck.addSlide({ elements: [{
    type: 'group', name: 'Unused fallback', accessibility: { description: 'Sales & "orders"' }, box, coordinateSize: box,
    action: { type: 'url', url: 'https://example.com/report', tooltip: 'Open sales report' },
    children: [inner, { ...inner, name: 'Decorative group', accessibility: { decorative: true } }],
  }] });
  const svg = (await renderPresentationToSvg(deck)).slides[0].svg;
  const browser = await chromium.launch();
  t.after(() => browser.close());
  const page = await browser.newPage(), cdp = await page.context().newCDPSession(page);
  for (const standalone of [false, true]) {
    if (standalone) await page.goto(`data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`);
    else await page.setContent(svg);
    const { nodes } = await cdp.send('Accessibility.getFullAXTree');
    const byId = new Map(nodes.map(node => [node.nodeId, node]));
    const named = (name, role) => {
      const matches = nodes.filter(node => !node.ignored && node.name?.value === name && !['StaticText', 'InlineTextBox'].includes(node.role?.value));
      assert.equal(matches.length, 1, name);
      assert.equal(matches[0].role.value, role, name);
      return matches[0];
    };
    const closestNamedAncestor = node => {
      for (let parent = byId.get(node.parentId); parent; parent = byId.get(parent.parentId)) {
        if (!parent.ignored && ['link', 'group', 'table'].includes(parent.role?.value) && parent.name?.value) return parent;
      }
      return undefined;
    };
    const link = named('Open sales report', 'link'), outer = named('Sales & "orders"', 'group'), group = named('Child group', 'group');
    const tableNode = named('Orders table', 'table'), chartNode = named('Orders chart', 'image'), header = named('Quantity', 'columnheader');
    for (const [node, parent] of [[outer, link], [group, outer], [tableNode, group], [chartNode, group], [header, tableNode]]) {
      assert.equal(closestNamedAncestor(node)?.nodeId, parent.nodeId, node.name.value);
    }
    assert.ok(!nodes.some(node => !node.ignored && ['Decorative group', 'Unused fallback'].includes(node.name?.value)));
  }
});

test('Tab skips links inside decorative subtrees without removing the object link', async t => {
  const deck = createPresentation(), box = { x: 20, y: 20, width: 200, height: 100 };
  const action = slug => ({ type: 'url', url: `https://example.com/${slug}`, tooltip: slug });
  const shape = slug => ({ type: 'shape', shape: 'rect', box, action: action(slug) });
  const text = slug => [{ runs: [{ text: slug, action: action(slug) }] }];
  const group = (children, extra = {}) => ({ type: 'group', box, coordinateSize: box, children, ...extra });
  deck.addSlide({ elements: [
    shape('before'),
    group([
      group([shape('hidden-object'), { type: 'text', box, content: text('hidden-text') }]),
      { type: 'shape', shape: 'rect', box, text: { content: text('hidden-shape-text') } },
      { type: 'table', box, columns: [200], rows: [{ cells: [{ content: text('hidden-cell') }] }] },
    ], { accessibility: { decorative: true }, action: action('outer') }),
    { type: 'text', box, content: text('decorative-text'), accessibility: { decorative: true }, action: action('decorative-frame') },
    { type: 'text', box, content: text('visible-text') },
    shape('after'),
  ] });
  const svg = (await renderPresentationToSvg(deck)).slides[0].svg;
  const browser = await chromium.launch();
  t.after(() => browser.close());
  const page = await browser.newPage();
  const readFocus = () => page.evaluate(() => document.activeElement?.getAttribute('href')?.split('/').at(-1));
  for (const standalone of [false, true]) {
    if (standalone) await page.goto(`data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`);
    else await page.setContent(svg);
    const focusOrder = [];
    for (let i = 0; i < 5; i++) {
      await page.keyboard.press('Tab');
      focusOrder.push(await readFocus());
    }
    assert.deepEqual(focusOrder, ['before', 'outer', 'decorative-frame', 'visible-text', 'after']);
    const reverse = [];
    for (let i = 0; i < 4; i++) {
      await page.keyboard.press('Shift+Tab');
      reverse.push(await readFocus());
    }
    assert.deepEqual(reverse, ['visible-text', 'decorative-frame', 'outer', 'before']);
    for (const slug of ['hidden-object', 'hidden-text', 'hidden-shape-text', 'hidden-cell', 'decorative-text']) {
      const links = page.locator(`a[href="https://example.com/${slug}"]`);
      assert.equal(await links.count(), 1, slug);
      assert.equal(await links.getAttribute('tabindex'), '-1', slug);
    }
  }
});

test('explicit SVG reading order keeps pixels, group descendants and inherited object positions', async t => {
  const deck = createPresentation({ size: { width: 360, height: 200 } }), box = { x: 30, y: 25, width: 130, height: 80 };
  const shape = (id, name, color) => ({ type: 'shape', shape: 'rect', id, name, box, style: { fill: { type: 'solid', color } } });
  deck.defineSlideLayout({ id: 'template', name: 'Template', elements: [shape('shared', 'Shared shape', '#333333')] });
  const slide = deck.addSlide({ id: 'orders', layoutId: 'template', elements: [
    shape('a <&>', 'First shape', '#ff0000'),
    { type: 'group', id: 'group', name: 'Table group', box, coordinateSize: box, children: [
      { type: 'table', id: 'child', name: 'Orders table', box, columns: [130], rows: [{ cells: [{ content: 'Units', headerScope: 'column' }] }] },
    ] },
    { type: 'chart', chartType: 'bar', id: 'chart', name: 'Sales chart', box, categories: ['A'], series: [{ name: 'S', values: [1] }] },
    { ...shape('decoration', 'Decorative shape', '#444444'), accessibility: { decorative: true } },
    { ...shape('hidden', 'Hidden shape', '#555555'), hidden: true },
  ] });
  deck.addSlide({ id: 'second', elements: [shape('other', 'Other shape', '#666666')] });
  const order = ['chart', 'group', 'a <&>'];
  const options = { idNamespace: 'preview test', readingOrder: [{ slideId: slide.id, elementIds: order }] };
  const baseline = await renderPresentationToSvg(deck, { idNamespace: options.idNamespace });
  const ordered = await renderPresentationToSvg(deck, options);
  assert.equal(ordered.slides[1].svg, baseline.slides[1].svg);
  const browser = await chromium.launch();
  t.after(() => browser.close());
  const page = await browser.newPage(), cdp = await page.context().newCDPSession(page);
  const names = new Set(['Shared shape', 'First shape', 'Table group', 'Orders table', 'Sales chart', 'Units', 'Decorative shape', 'Hidden shape']);
  async function observed(svg, standalone = false) {
    if (standalone) await page.goto(`data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`);
    else await page.setContent(svg);
    const { nodes } = await cdp.send('Accessibility.getFullAXTree'), byId = new Map(nodes.map(node => [node.nodeId, node]));
    const result = [];
    const visit = id => {
      const node = byId.get(id);
      if (!node) return;
      if (!node.ignored && names.has(node.name?.value) && ['image', 'group', 'table', 'columnheader'].includes(node.role?.value)) result.push(node.name.value);
      for (const child of node.childIds ?? []) visit(child);
    };
    visit(nodes[0].nodeId);
    return result;
  }
  assert.deepEqual(await observed(baseline.slides[0].svg), ['Shared shape', 'First shape', 'Table group', 'Orders table', 'Units', 'Sales chart']);
  const screenshot = await page.locator('svg').first().screenshot();
  assert.deepEqual(await observed(ordered.slides[0].svg), ['Shared shape', 'Sales chart', 'Table group', 'Orders table', 'Units', 'First shape']);
  assert.deepEqual(await page.locator('svg').first().screenshot(), screenshot);
  assert.deepEqual(await observed(ordered.slides[0].svg, true), ['Shared shape', 'Sales chart', 'Table group', 'Orders table', 'Units', 'First shape']);
  const separate = await renderPresentationToSvg(deck, { ...options, idNamespace: 'other preview' });
  await page.goto('about:blank');
  await page.setContent(ordered.slides[0].svg + separate.slides[0].svg);
  const idsAreLocal = await page.locator('svg[aria-owns]').evaluateAll(slides => slides.every(svg =>
    svg.getAttribute('aria-owns').split(' ').every(id => svg.contains(document.getElementById(id)))));
  assert.equal(idsAreLocal, true);
  const layout = resolveLayout(deck), first = await renderIncrementalSvg(layout, { ...options, environmentRevision: '1' });
  assert.deepEqual(first.result, ordered);
  const alternate = { ...options, environmentRevision: '1', readingOrder: [{ slideId: slide.id, elementIds: [...order].reverse() }] };
  const second = await renderIncrementalSvg(layout, alternate, first.cache);
  assert.deepEqual(second.renderedSlideIds, ['orders']);
  assert.deepEqual(second.result, await renderLayoutToSvg(layout, alternate));
  for (const ids of [['chart', 'group'], ['chart', 'group', 'chart'], ['chart', 'group', 'child'], ['chart', 'group', 'decoration'], ['chart', 'group', 'hidden'], ['chart', 'group', 'shared']]) {
    await assert.rejects(renderLayoutToSvg(layout, { readingOrder: [{ slideId: slide.id, elementIds: ids }] }), /slide-owned root/);
  }
  await assert.rejects(renderLayoutToSvg(layout, { readingOrder: [options.readingOrder[0], options.readingOrder[0]] }), /distinct slide IDs/);
  await assert.rejects(renderLayoutToSvg(layout, { readingOrder: [{ slideId: 'missing', elementIds: order }] }), /existing visible slide/);
  const linked = structuredClone(slide.elements.find(element => element.id === 'group'));
  linked.children[0].rows[0].cells[0].content = [{ runs: [{ text: 'Units', action: { type: 'url', url: 'https://example.com' } }] }];
  slide.replaceElement(linked.id, linked);
  await assert.rejects(renderPresentationToSvg(deck, options), /keyboard focus order/);
  slide.removeElement('chart');
  await assert.rejects(renderPresentationToSvg(deck, options), /slide-owned root/);
});

test('nested reading orders compose with the slide order without moving objects across groups', async t => {
  const deck = createPresentation({ size: { width: 300, height: 160 } }), box = { x: 10, y: 10, width: 130, height: 60 };
  const shape = (id, color = '#001122') => ({ type: 'shape', shape: 'rect', id, name: id, box, style: { fill: { type: 'solid', color } } });
  const group = (id, children, extra = {}) => ({ type: 'group', id, name: id, box, coordinateSize: box, children, ...extra });
  deck.defineSlideLayout({ id: 'layout', name: 'Layout', elements: [group('inherited', [shape('inherited-child')])] });
  const slide = deck.addSlide({ id: 'nested', layoutId: 'layout', elements: [
    group('outer', [shape('first'), group('inner', [shape('inside-a', '#004488'), shape('inside-b', '#ff8800')]), shape('last')]),
    shape('root'),
    group('decorated', [group('decorated-inner', [shape('not-readable')])], { accessibility: { decorative: true } }),
    group('hidden-group', [group('hidden-inner', [shape('not-visible')])], { hidden: true }),
  ] });
  deck.addSlide({ id: 'other', elements: [shape('other-root')] });
  const layout = resolveLayout(deck), environmentRevision = 'nested-orders';
  const root = { slideId: 'nested', elementIds: ['root', 'outer'] };
  const outer = { slideId: 'nested', groupId: 'outer', elementIds: ['last', 'inner', 'first'] };
  const inner = { slideId: 'nested', groupId: 'inner', elementIds: ['inside-b', 'inside-a'] };
  const browser = await chromium.launch(); t.after(() => browser.close());
  const page = await browser.newPage(), cdp = await page.context().newCDPSession(page);
  const names = new Set(['inherited', 'inherited-child', 'root', 'outer', 'first', 'inner', 'inside-a', 'inside-b', 'last', 'not-readable', 'not-visible']);
  async function observe(svg) {
    await page.setContent(svg);
    const { nodes } = await cdp.send('Accessibility.getFullAXTree'), index = new Map(nodes.map(node => [node.nodeId, node]));
    const result = [];
    const visit = (id, ancestor) => {
      const node = index.get(id);
      if (!node) return;
      const named = !node.ignored && names.has(node.name?.value) && ['image', 'group'].includes(node.role?.value);
      const current = named ? node.name.value : ancestor;
      if (named) result.push([current, ancestor]);
      for (const child of node.childIds ?? []) visit(child, current);
    };
    visit(nodes[0].nodeId, null);
    return result;
  }
  const baseline = await renderLayoutToSvg(layout);
  assert.deepEqual(await observe(baseline.slides[0].svg), [
    ['inherited', null], ['inherited-child', 'inherited'], ['outer', null], ['first', 'outer'], ['inner', 'outer'],
    ['inside-a', 'inner'], ['inside-b', 'inner'], ['last', 'outer'], ['root', null],
  ]);
  const pixels = await page.locator('svg').screenshot();
  const options = { readingOrder: [inner, root, outer], environmentRevision };
  const ordered = await renderIncrementalSvg(layout, options);
  assert.deepEqual(await observe(ordered.result.slides[0].svg), [
    ['inherited', null], ['inherited-child', 'inherited'], ['root', null], ['outer', null], ['last', 'outer'],
    ['inner', 'outer'], ['inside-b', 'inner'], ['inside-a', 'inner'], ['first', 'outer'],
  ]);
  assert.deepEqual(await page.locator('svg').screenshot(), pixels);
  assert.deepEqual((await renderIncrementalSvg(layout, { ...options, readingOrder: [root, outer, inner] }, ordered.cache)).renderedSlideIds, []);
  assert.deepEqual((await renderIncrementalSvg(layout, { ...options, readingOrder: [root, outer] }, ordered.cache)).renderedSlideIds, ['nested']);
  const onlyGroup = await renderLayoutToSvg(layout, { readingOrder: [inner] });
  assert.deepEqual((await observe(onlyGroup.slides[0].svg)).map(([name]) => name), ['inherited', 'inherited-child', 'outer', 'first', 'inner', 'inside-b', 'inside-a', 'last', 'root']);
  for (const groupId of ['inherited', 'decorated-inner', 'hidden-inner', 'root']) {
    await assert.rejects(renderLayoutToSvg(layout, { readingOrder: [{ ...inner, groupId }] }), /slide-owned group/);
  }
  await assert.rejects(renderLayoutToSvg(layout, { readingOrder: [inner, inner] }), /distinct group IDs/);
  await assert.rejects(renderLayoutToSvg(layout, { readingOrder: [{ ...outer, elementIds: ['first', 'last', 'inside-a'] }] }), /direct group child/);
  const clone = slide.duplicateElement('outer');
  assert.notEqual(clone.id, 'outer');
  await assert.rejects(renderPresentationToSvg(deck, options), /slide-owned root/);
  assert.deepEqual((await renderPresentationToSvg(deck, { readingOrder: [inner] })).slides.length, 2);
  slide.removeElement('outer');
  await assert.rejects(renderPresentationToSvg(deck, { readingOrder: [inner] }), /slide-owned group/);
});

test('authored reading order flows into SVG defaults but leaves native drawing bytes unchanged', async () => {
  const deck = createPresentation(), box = { x: 10, y: 20, width: 100, height: 45 };
  const shape = id => ({ id, type: 'shape', name: id, shape: 'rect', box });
  const slide = deck.addSlide({ id: 'authored', elements: [shape('first'), shape('second')], readingOrder: [{ elementIds: ['second', 'first'] }] });
  const svg = await renderPresentationToSvg(deck);
  assert.match(svg.slides[0].svg, /aria-owns=/);
  assert.deepEqual(svg, await renderPresentationToSvg(deck, { readingOrder: [{ slideId: 'authored', elementIds: ['second', 'first'] }] }));
  const override = await renderPresentationToSvg(deck, { readingOrder: [] });
  assert.doesNotMatch(override.slides[0].svg, /aria-owns=/);
  const layout = resolveLayout(deck), incremental = await renderIncrementalSvg(layout, { environmentRevision: 'authored' });
  assert.deepEqual(incremental.result, svg);
  const cleared = await renderIncrementalSvg(layout, { environmentRevision: 'authored', readingOrder: [] }, incremental.cache);
  assert.deepEqual(cleared.renderedSlideIds, ['authored']);
  assert.deepEqual(cleared.result, override);
  const native = await generatePptx(deck);
  assert.deepEqual(native.warnings.filter(w => w.code === 'reading-order-not-exported').map(w => [w.slideId, w.path]), [['authored', 'slides.0.readingOrder']]);
  slide.setHidden(true);
  assert.doesNotMatch((await renderPresentationToSvg(deck)).slides[0].svg, /aria-owns=/);
  slide.setHidden(false);
  slide.setReadingOrder();
  assert.deepEqual((await generatePptx(deck)).bytes, native.bytes);
});
