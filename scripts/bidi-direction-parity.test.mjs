import test from 'node:test';
import assert from 'node:assert/strict';
import { createPresentation, normalizePresentation, validatePresentation } from '../packages/core/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';
import { renderPresentationToSvg } from '../packages/svg-renderer/dist/index.js';
import { importPptx } from '../packages/pptx-parser/dist/index.js';
import { readGeneratedZip } from '../tests/support/read-generated-zip.mjs';

const paragraphs = [
  { style: { direction: 'rtl', align: 'left' }, runs: [{ text: 'שלום (ABC 123) סוף' }] },
  { style: { direction: 'ltr', align: 'left' }, runs: [{ text: 'ABC (שלום 123) XYZ' }] },
  { runs: [{ text: 'Plain' }] },
];
function fixture() {
  const deck = createPresentation(), slide = deck.addSlide({ id: 'direction-slide', notes: paragraphs });
  slide.addElement({ type: 'text', id: 'direction-text', box: { x: 20, y: 20, width: 300, height: 180 }, frame: { wrap: false }, content: paragraphs });
  slide.addElement({ type: 'shape', id: 'direction-shape', shape: 'rect', box: { x: 350, y: 20, width: 300, height: 180 }, text: { content: paragraphs } });
  slide.addElement({ type: 'table', id: 'direction-table', box: { x: 20, y: 220, width: 300, height: 180 }, columns: [300], rows: [{ height: 180, cells: [{ content: paragraphs }] }] });
  return deck;
}

test('paragraph direction authors, projects and preserves source while style edits remain blocked', async () => {
  const deck = fixture(), normalized = normalizePresentation(deck);
  assert.deepEqual(validatePresentation(deck), []);
  assert.deepEqual(normalized.slides[0].elements[0].content.map(p => p.style.direction), ['rtl', 'ltr', undefined]);
  const exported = await generatePptx(deck), parts = readGeneratedZip(exported.bytes);
  const slide = parts.get('ppt/slides/slide1.xml').toString(), notes = parts.get('ppt/notesSlides/notesSlide1.xml').toString();
  assert.equal((slide.match(/<a:pPr[^>]*\brtl="1"/g) ?? []).length, 3);
  assert.equal((slide.match(/<a:pPr[^>]*\brtl="0"/g) ?? []).length, 3);
  assert.equal((notes.match(/<a:pPr[^>]*\brtl="1"/g) ?? []).length, 1);
  assert.equal((notes.match(/<a:pPr[^>]*\brtl="0"/g) ?? []).length, 1);
  const preview = (await renderPresentationToSvg(deck)).slides[0].svg;
  assert.match(preview, /direction:rtl/); assert.match(preview, /direction:ltr/);
  const session = await importPptx(exported.bytes, { projection: 'content', editSaveProfile: 'text-runs-v3' });
  try {
    const items = session.document.slides[0].elements;
    const directions = items.map(item => item.type === 'shape' ? item.text.content.map(p => p.style.direction) : item.type === 'table' ? item.rows[0].cells[0].content.map(p => p.style.direction) : item.content.map(p => p.style.direction));
    assert.deepEqual(directions, Array(3).fill(['rtl', 'ltr', undefined]));
    assert.deepEqual(session.document.slides[0].notes.map(p => p.style.direction), ['rtl', 'ltr', undefined]);
    const noOp = await session.prepareSave({ mode: 'no-op' }); assert.equal(noOp.status, 'ready'); assert.deepEqual((await session.save(noOp.plan)).bytes, exported.bytes);
    const slideDoc = session.document.slides[0], original = items[0], changed = structuredClone(original);
    changed.content[0].style.direction = 'ltr'; slideDoc.replaceElement(original.id, changed);
    assert.equal((await session.prepareSave({ mode: 'edit' })).status, 'blocked');
  } finally { session.dispose(); }
});

test('invalid direction produces a located validation error', () => {
  const deck = fixture(), slide = deck.slides[0], item = structuredClone(slide.elements[0]);
  item.content[0].style.direction = 'auto'; slide.replaceElement(item.id, item);
  assert.ok(validatePresentation(deck).some(d => d.code === 'invalid-paragraph-direction' && d.path.endsWith('.direction')));
});

test('explicit paragraph direction survives three fresh export/import rounds', async () => {
  let bytes = (await generatePptx(fixture())).bytes;
  for (let round = 0; round < 3; round++) {
    const session = await importPptx(bytes, { projection: 'content', editSaveProfile: 'text-runs-v3' });
    try {
      const first = session.document.slides[0].elements[0];
      assert.deepEqual(first.content.map(paragraph => paragraph.style.direction), ['rtl', 'ltr', undefined]);
      const noOp = await session.prepareSave({ mode: 'no-op' });
      assert.equal(noOp.status, 'ready'); assert.deepEqual((await session.save(noOp.plan)).bytes, bytes);
      bytes = (await generatePptx(session.document)).bytes;
      const xml = readGeneratedZip(bytes).get('ppt/slides/slide1.xml').toString();
      assert.equal((xml.match(/<a:pPr[^>]*\brtl="1"/g) ?? []).length, 3);
      assert.equal((xml.match(/<a:pPr[^>]*\brtl="0"/g) ?? []).length, 3);
    } finally { session.dispose(); }
  }
});

test('direction inherits from presets and tables and local declarations override it', () => {
  const deck = createPresentation({ textStylePresets: { rtl: { paragraph: { direction: 'rtl' } } } });
  const slide = deck.addSlide();
  slide.addElement({ type: 'text', id: 'preset', box: { x: 0, y: 0, width: 200, height: 80 }, textStylePreset: 'rtl', content: [
    { runs: [{ text: 'one' }] }, { style: { direction: 'ltr' }, runs: [{ text: 'two' }] },
  ] });
  slide.addElement({ type: 'table', id: 'table', box: { x: 0, y: 100, width: 200, height: 80 }, columns: [200],
    style: { wholeTable: { text: { paragraph: { direction: 'rtl' } } } },
    rows: [{ cells: [{ content: [{ runs: [{ text: 'one' }] }, { style: { direction: 'ltr' }, runs: [{ text: 'two' }] }] }] }] });
  assert.deepEqual(validatePresentation(deck), []);
  const [text, table] = normalizePresentation(deck).slides[0].elements;
  assert.deepEqual(text.content.map(p => p.style.direction), ['rtl', 'ltr']);
  assert.deepEqual(table.rows[0].cells[0].content.map(p => p.style.direction), ['rtl', 'ltr']);
});

test('direction preset reset can preserve native placeholder inheritance', async () => {
  const deck = createPresentation({ textStylePresets: { ltr: { paragraph: { direction: 'ltr' } } } });
  deck.defineSlideLayout({ id: 'layout', name: 'Layout', placeholders: [{ key: 'body', kind: 'body', box: { x: 0, y: 0, width: 300, height: 80 }, textStyle: { paragraph: { direction: 'rtl' } } }] });
  deck.addSlide({ layoutId: 'layout' }).addElement({ type: 'text', id: 'body', placeholderKey: 'body', textStylePreset: 'ltr', textStylePresetReset: ['paragraph.direction'], content: 'Hello' });
  assert.deepEqual(validatePresentation(deck), []);
  assert.equal(normalizePresentation(deck).slides[0].elements[0].content[0].style.direction, 'rtl');
  const output = await generatePptx(deck, { nativePlaceholderText: 'inherit' }), parts = readGeneratedZip(output.bytes);
  assert.doesNotMatch(parts.get('ppt/slides/slide1.xml').toString(), /<a:pPr[^>]*\brtl=/);
  const layouts = [...parts].filter(([name]) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(name)).map(([, bytes]) => bytes.toString());
  assert.ok(layouts.some(xml => /<a:lvl1pPr[^>]*\brtl="1"/.test(xml)));
  assert.ok(!output.warnings.some(w => w.code === 'native-text-inheritance-materialized' && w.elementId === 'body'));
  const imported = await importPptx(output.bytes, { projection: 'content' });
  try {
    assert.equal(imported.document.slides[0].elements[0].content[0].style.direction, 'rtl');
    assert.ok(imported.report.textStyleInheritance.properties.some(item => item.property === 'paragraph.direction' && item.value === 'rtl' && item.level === 'layout-placeholder'));
  } finally { imported.dispose(); }
});
