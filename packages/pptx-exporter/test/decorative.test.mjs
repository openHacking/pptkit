import test from 'node:test';
import assert from 'node:assert/strict';
import { createPresentation } from '@pptkit/core';
import { generatePptx } from '../dist/index.js';
import { readGeneratedZip } from '../../../tests/support/read-generated-zip.mjs';
const png = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAgAAAAECAIAAAA8r+mnAAAAIUlEQVR4nGNQCb8GRze90uCIAaeE0tE4OPp1RhSOcEoAAP18L9GcKmUbAAAAAElFTkSuQmCC', 'base64');

test('native decorative extension is attached to selected object identities without losing actions', async () => {
  const box = { x: 0, y: 0, width: 100, height: 50 };
  const flag = { decorative: true, description: 'Description & context' };
  const deck = createPresentation({ masters: [{ id: 'master', name: 'Master', elements: [{ id: 'master-shape', name: 'Master shape', type: 'shape', shape: 'rect', box, accessibility: flag }] }] });
  deck.registerAsset({ id: 'asset', kind: 'image', source: { type: 'bytes', value: png.toString('base64'), encoding: 'base64' }, mimeType: 'image/png', width: 8, height: 4 });
  deck.defineSlideLayout({ id: 'layout', name: 'Layout', masterId: 'master', elements: [{ id: 'layout-shape', name: 'Layout shape', type: 'shape', shape: 'rect', box, accessibility: flag }] });
  deck.addSlide({ id: 'page', layoutId: 'layout', elements: [
    { id: 'text', name: 'Text', type: 'text', box, content: 'A', accessibility: flag, action: { type: 'url', url: 'https://example.com' } },
    { id: 'shape', name: 'Shape', type: 'shape', shape: 'rect', box, accessibility: flag },
    { id: 'image', name: 'Image', type: 'image', assetId: 'asset', box, accessibility: flag },
    { id: 'line', name: 'Line', type: 'connector', start: { x: 1, y: 1 }, end: { x: 10, y: 10 }, accessibility: flag },
    { id: 'chart', name: 'Chart', type: 'chart', box, chartType: 'bar', categories: ['A'], series: [{ name: 'S', values: [1] }], accessibility: flag },
    { id: 'table', name: 'Table', type: 'table', box, columns: [100], rows: [{ cells: [{ content: 'A' }] }], accessibility: flag, transform: { rotation: 15 } },
    { id: 'group', name: 'Group', type: 'group', box, coordinateSize: { width: 100, height: 50 }, accessibility: flag, children: [{ id: 'child', name: 'Child', type: 'shape', shape: 'rect', box }] },
    { id: 'ordinary', name: 'Ordinary', type: 'shape', shape: 'rect', box, accessibility: { decorative: false } },
  ] });
  const result = await generatePptx(deck), parts = readGeneratedZip(result.bytes);
  const slide = parts.get('ppt/slides/slide1.xml').toString();
  const owners = [...slide.matchAll(/<p:cNvPr\b([^>]*?)(?<!\/)>([\s\S]*?)<\/p:cNvPr>/g)].map(([, attributes, children]) => ({ name: attributes.match(/\bname="([^"]*)"/)?.[1], attributes, children }));
  const marked = owners.filter(owner => owner.children.includes('adec:decorative'));
  assert.deepEqual(marked.map(owner => owner.name), ['Text', 'Shape', 'Image', 'Line', 'Chart', 'Table', 'Table', 'Group']);
  for (const owner of marked) {
    assert.match(owner.children, /<a:extLst><a:ext uri="\{C183D7F6-B498-43B3-948B-1728B52AA6E4\}"><adec:decorative xmlns:adec="http:\/\/schemas\.microsoft\.com\/office\/drawing\/2017\/decorative" val="1"\/><\/a:ext><\/a:extLst>/);
    assert.match(owner.attributes, /descr="Description &amp; context"/);
  }
  assert.match(owners.find(owner => owner.name === 'Text').children, /<a:hlinkClick[^>]+\/>.*<a:extLst>/);
  assert.doesNotMatch(owners.find(owner => owner.name === 'Ordinary').children, /decorative/);
  assert.doesNotMatch(owners.find(owner => owner.name === 'Child').children, /decorative/);
  const layout = [...parts].find(([name]) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(name) && parts.get(name).toString().includes('Layout shape'))[1].toString();
  const master = [...parts].find(([name]) => /^ppt\/slideMasters\/slideMaster\d+\.xml$/.test(name) && parts.get(name).toString().includes('Master shape'))[1].toString();
  assert.match(layout, /name="Layout shape"[^>]*>\s*<a:extLst>.*<adec:decorative/);
  assert.match(master, /name="Master shape"[^>]*>\s*<a:extLst>.*<adec:decorative/);
});

test("explicit native table presets report that direct cell formatting can mask the preset", async () => {
  const deck = createPresentation();
  deck.addSlide({ id: "page", elements: [{
    id: "preset-table", type: "table", nativeTableStylePreset: "medium-2-accent-2",
    box: { x: 0, y: 0, width: 100, height: 50 }, columns: [100],
    rows: [{ cells: [{ content: "Cell" }] }],
  }] });
  const result = await generatePptx(deck);
  const warning = result.warnings.find(item => item.code === "native-table-preset-direct-formatting");
  assert.equal(warning?.slideId, "page");
  assert.equal(warning?.elementId, "preset-table");
});
