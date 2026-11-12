import test from 'node:test';
import assert from 'node:assert/strict';
import { createPresentation, normalizePresentation, resolveTextFontSegments, validatePresentation } from '../packages/core/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';
import { renderPresentationToSvg } from '../packages/svg-renderer/dist/index.js';
import { importPptx } from '../packages/pptx-parser/dist/index.js';
import { readGeneratedZip } from '../tests/support/read-generated-zip.mjs';

function fixture() {
  const deck = createPresentation(), slide = deck.addSlide({ id: 'slide' });
  const runs = [
    { text: 'ABC ', style: { fontFamily: 'Latin Face' } },
    { text: 'مرحبا', style: { fontFamily: 'Latin Face', complexScriptFontFamily: 'Complex Face', rightToLeft: true } },
    { text: ' - 123 ', style: { fontFamily: 'Latin Face', complexScriptFontFamily: 'Complex Face', rightToLeft: true } },
    { text: 'XYZ', style: { fontFamily: 'Latin Face', rightToLeft: false } },
  ];
  slide.addElement({ type: 'text', id: 'run-rtl', box: { x: 0, y: 0, width: 400, height: 120 }, content: [{ style: { direction: 'ltr' }, runs }] });
  return deck;
}

test('run RTL maps explicit true and false, round-trips and reports preview limits', async () => {
  const deck = fixture(), before = normalizePresentation(deck), runs = before.slides[0].elements[0].content[0].runs;
  assert.deepEqual(validatePresentation(deck), []);
  assert.deepEqual(runs.map(run => run.style.rightToLeft), [undefined, true, true, false]);
  assert.ok(resolveTextFontSegments(' - 123 ', runs[2].style, before.theme).every(segment => segment.family === 'Complex Face'));
  assert.ok(resolveTextFontSegments(' - 123 ', { ...runs[2].style, rightToLeft: false }, before.theme).every(segment => segment.family === 'Latin Face'));
  assert.equal(resolveTextFontSegments('A', { ...runs[2].style, scriptFontFamilies: { Latn: 'Mapped Face' } }, before.theme)[0].family, 'Mapped Face');
  const preview = await renderPresentationToSvg(deck);
  const warnings = preview.warnings.filter(warning => warning.code === 'text-run-rtl-approximation' && warning.elementId === 'run-rtl');
  assert.equal(warnings.length, 1);
  assert.match(warnings[0].message, /font selection may differ from native applications/);
  assert.match(preview.slides[0].svg, /font-family:&quot;Complex Face&quot;/);
  let bytes = (await generatePptx(deck)).bytes;
  for (let round = 0; round < 3; round++) {
    const xml = readGeneratedZip(bytes).get('ppt/slides/slide1.xml').toString();
    assert.equal((xml.match(/<a:rtl val="1"\/>/g) ?? []).length, 2);
    assert.equal((xml.match(/<a:rtl val="0"\/>/g) ?? []).length, 1);
    const session = await importPptx(bytes, { projection: 'content', textFontProfile: 'script-families-v1', editSaveProfile: 'text-runs-v3' });
    try {
      const slide = session.document.slides[0], original = slide.elements[0];
      assert.deepEqual(original.content[0].runs.map(run => run.style.rightToLeft), [undefined, true, true, false]);
      assert.equal(session.report.textStyleInheritance.properties.filter(item => item.property === 'run.rightToLeft' && item.status === 'mapped').length, 3);
      const noOp = await session.prepareSave({ mode: 'no-op' }); assert.equal(noOp.status, 'ready'); assert.deepEqual((await session.save(noOp.plan)).bytes, bytes);
      const changed = structuredClone(original); changed.content[0].runs[1].style.rightToLeft = false; slide.replaceElement(original.id, changed);
      assert.equal((await session.prepareSave({ mode: 'edit' })).status, 'blocked'); slide.replaceElement(original.id, original);
      bytes = (await generatePptx(session.document)).bytes;
    } finally { session.dispose(); }
  }
  assert.deepEqual(normalizePresentation(deck), before);
});

test('run RTL inherits in presets and explicit false clears it; invalid values are rejected', () => {
  const deck = createPresentation({ textStylePresets: { rtl: { run: { rightToLeft: true } } } });
  const slide = deck.addSlide(); slide.addElement({ type: 'text', id: 'text', textStylePreset: 'rtl', box: { x: 0, y: 0, width: 200, height: 60 }, content: [{ runs: [{ text: 'one' }, { text: 'two', style: { rightToLeft: false } }] }] });
  assert.deepEqual(validatePresentation(deck), []);
  assert.deepEqual(normalizePresentation(deck).slides[0].elements[0].content[0].runs.map(run => run.style.rightToLeft), [true, false]);
  const changed = structuredClone(slide.elements[0]); changed.content[0].runs[0].style = { rightToLeft: 'yes' }; slide.replaceElement(changed.id, changed);
  assert.ok(validatePresentation(deck).some(item => item.code === 'invalid-run-direction' && item.path.endsWith('.rightToLeft')));
});

test('run RTL carries through shape, table and notes while preview warnings stay per owner', async () => {
  const paragraphs = [{ runs: [{ text: 'مرحبا', style: { rightToLeft: true } }, { text: '123', style: { rightToLeft: false } }] }];
  const deck = createPresentation(), slide = deck.addSlide({ id: 'owners', notes: paragraphs });
  slide.addElement({ type: 'shape', id: 'shape', shape: 'rect', box: { x: 0, y: 0, width: 200, height: 80 }, text: { content: paragraphs } });
  slide.addElement({ type: 'table', id: 'table', box: { x: 0, y: 100, width: 200, height: 80 }, columns: [100, 100], rows: [{ cells: [{ content: paragraphs }, { content: paragraphs }] }] });
  assert.deepEqual(validatePresentation(deck), []);
  const preview = await renderPresentationToSvg(deck);
  assert.deepEqual(preview.warnings.filter(w => w.code === 'text-run-rtl-approximation').map(w => w.elementId).sort(), ['shape', 'table']);
  const output = await generatePptx(deck), parts = readGeneratedZip(output.bytes);
  assert.equal((parts.get('ppt/slides/slide1.xml').toString().match(/<a:rtl val="1"\/>/g) ?? []).length, 3);
  assert.equal((parts.get('ppt/notesSlides/notesSlide1.xml').toString().match(/<a:rtl val="1"\/>/g) ?? []).length, 1);
  const session = await importPptx(output.bytes, { projection: 'content', textFontProfile: 'script-families-v1' });
  try {
    const [shape, table] = session.document.slides[0].elements;
    for (const content of [shape.text.content, ...table.rows[0].cells.map(cell => cell.content), session.document.slides[0].notes]) {
      assert.deepEqual(content[0].runs.map(run => run.style.rightToLeft), [true, false]);
    }
  } finally { session.dispose(); }
});

test('native layout run RTL defaults preserve explicit opposite local overrides', async () => {
  for (const direction of [true, false]) {
    const deck = createPresentation({ textStylePresets: { preset: { run: { rightToLeft: !direction } } } });
    deck.defineSlideLayout({ id: 'layout', name: 'Layout', placeholders: [{ key: 'body', kind: 'body', box: { x: 0, y: 0, width: 300, height: 80 }, textStyle: { run: { rightToLeft: direction } } }] });
    deck.addSlide({ layoutId: 'layout' }).addElement({ type: 'text', id: 'body', placeholderKey: 'body', textStylePreset: 'preset', textStylePresetReset: ['rightToLeft'], content: [{ runs: [{ text: 'مرحبا' }, { text: '123', style: { rightToLeft: !direction } }] }] });
    deck.addSlide({ layoutId: 'layout' }).addElement({ type: 'text', id: 'body-shared', placeholderKey: 'body', content: [{ runs: [{ text: 'مرحبا' }, { text: '123', style: { rightToLeft: !direction } }] }] });
    assert.deepEqual(validatePresentation(deck), []);
    const materialized = readGeneratedZip((await generatePptx(deck)).bytes).get('ppt/slides/slide1.xml').toString();
    assert.match(materialized, new RegExp(`<a:rtl val="${Number(direction)}"/>`));
    const output = await generatePptx(deck, { nativePlaceholderText: 'inherit' }), parts = readGeneratedZip(output.bytes);
    const slide = parts.get('ppt/slides/slide1.xml').toString();
    assert.doesNotMatch(slide, new RegExp(`<a:rtl val="${Number(direction)}"/>`));
    assert.match(slide, new RegExp(`<a:rtl val="${Number(!direction)}"/>`));
    const layouts = [...parts].filter(([name]) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(name)).map(([, bytes]) => bytes.toString());
    assert.ok(layouts.some(xml => new RegExp(`<a:defRPr[^>]*><a:rtl val="${Number(direction)}"/></a:defRPr>`).test(xml)));
    assert.ok(!output.warnings.some(w => w.code === 'native-text-inheritance-materialized'));
    const imported = await importPptx(output.bytes, { projection: 'content' });
    try {
      for (const importedSlide of imported.document.slides) assert.deepEqual(importedSlide.elements[0].content[0].runs.map(run => run.style.rightToLeft), [direction, !direction]);
      assert.ok(imported.report.textStyleInheritance.properties.some(item => item.property === 'run.rightToLeft' && item.value === direction && item.level === 'layout-placeholder'));
    } finally { imported.dispose(); }
  }
});
