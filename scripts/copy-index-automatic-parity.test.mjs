import assert from 'node:assert/strict';
import test from 'node:test';
import { createPresentation, normalizePresentation } from '../packages/core/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';
import { readGeneratedZip } from '../tests/support/read-generated-zip.mjs';
import { png } from '../packages/pptx-parser/test/image-fixture.mjs';

const box = { x: 60, y: 40, width: 400, height: 60 };
const body = { x: 60, y: 130, width: 400, height: 60 };
const assets = { deduplicate: 'content', external: 'reject' };

test('automatic index matching emits the same OOXML parts as explicit per-page mappings', async () => {
  const source = createPresentation();
  source.registerAsset({ id: 'picture', kind: 'image', mimeType: 'image/png', source: { type: 'bytes', value: png }, width: 8, height: 4 });
  source.defineSlideLayout({ id: 'source', name: 'Source', placeholders: [{ key: 'heading', kind: 'title', index: 0, box }, { key: 'body', kind: 'body', index: 4294967295, box: body }] });
  for (const [id, other] of [['a', 'b'], ['b', 'a']]) source.addSlide({ id, layoutId: 'source', notes: `Notes ${id}`, elements: [
    { id: `heading-${id}`, type: 'text', placeholderKey: 'heading', content: `Heading ${id}`, action: { type: 'slide', slideId: other } },
    { id: `body-${id}`, type: 'text', placeholderKey: 'body', content: `Body ${id}` },
    { id: `picture-${id}`, type: 'image', assetId: 'picture', box: { x: 20, y: 220, width: 80, height: 40 } },
  ] });
  const makeDestination = automatic => {
    const target = createPresentation({ id: 'target' });
    for (const id of ['a', 'b']) target.defineSlideLayout({ id: `layout-${id}`, name: `Layout ${id}`, placeholders: [
      { key: `title-${id}`, kind: 'title', index: 0, box }, { key: `copy-${id}`, kind: 'body', index: 4294967295, box: body },
    ] });
    const options = { layoutId: null, layouts: { a: 'layout-a', b: 'layout-b' }, theme: 'destination', geometry: 'preserve', customData: 'reject', placeholderRoles: 'require-match-and-index', assets };
    const mapping = automatic ? { placeholderMatching: 'explicit-index' } : { perSlide: { a: { placeholders: { heading: 'title-a', body: 'copy-a' } }, b: { placeholders: { heading: 'title-b', body: 'copy-b' } } } };
    const result = target.copySlidesFrom(source, ['a', 'b'], { ...options, ...mapping });
    return { target, result };
  };
  const explicit = makeDestination(false), automatic = makeDestination(true);
  assert.deepEqual(automatic.result.slideIds, explicit.result.slideIds);
  assert.deepEqual(automatic.result.elementIds, explicit.result.elementIds);
  assert.deepEqual(automatic.result.assets, explicit.result.assets);
  assert.deepEqual(normalizePresentation(automatic.target), normalizePresentation(explicit.target));
  source.removeSlide('a'); source.removeSlide('b');
  const a = readGeneratedZip((await generatePptx(automatic.target)).bytes), b = readGeneratedZip((await generatePptx(explicit.target)).bytes);
  assert.deepEqual([...a.keys()].sort(), [...b.keys()].sort());
  for (const [part, bytes] of a) assert.deepEqual(bytes, b.get(part), part);
  for (const id of [1, 2]) {
    const xml = a.get(`ppt/slides/slide${id}.xml`).toString();
    assert.match(xml, /<p:ph type="title" idx="0"\/>/);
    assert.match(xml, /<p:ph type="body" idx="4294967295"\/>/);
  }
});
