import test from 'node:test';
import assert from 'node:assert/strict';
import { createPresentation, validatePresentation } from '../packages/core/dist/index.js';
import { resolveLayout } from '../packages/layout/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';
import { importPptx } from '../packages/pptx-parser/dist/index.js';
import { renderPresentationToSvg } from '../packages/svg-renderer/dist/index.js';
import { readGeneratedZip } from '../tests/support/read-generated-zip.mjs';

test('text frames preserve quarter turns independently of shape geometry and horizontal overrides', async () => {
  const deck = createPresentation();
  const slide = deck.addSlide();
  slide.addElement({ id: 'quarter', type: 'text', box: { x: 20, y: 30, width: 80, height: 150 }, frame: { textDirection: 'rotate90', margin: { left: 2, top: 3, right: 5, bottom: 7 } }, content: '日本語 ABC' });
  slide.addElement({ id: 'reverse', type: 'shape', shape: 'rect', box: { x: 130, y: 30, width: 80, height: 150 }, text: { frame: { textDirection: 'rotate270' }, content: 'ABC 123' } });
  slide.addElement({ id: 'horizontal', type: 'text', box: { x: 240, y: 30, width: 80, height: 150 }, frame: { textDirection: 'horizontal' }, content: 'ABC' });
  assert.deepEqual(validatePresentation(deck), []);
  const measured = [];
  const layout = resolveLayout(deck, { textMeasurement: { measure(request) {
    measured.push(request.availableWidth);
    return { bounds: { x: 0, y: 0, width: 20, height: 10 } };
  } } });
  assert.equal(measured[0], 140);
  assert.deepEqual([layout.analysis.measurements[0].bounds.width, layout.analysis.measurements[0].bounds.height], [10, 20]);
  assert.equal(layout.analysis.diagnostics.filter(d => d.code === 'text-frame-direction-measurement-approximate').length, 2);
  const bytes = (await generatePptx(deck)).bytes;
  const xml = readGeneratedZip(bytes).get('ppt/slides/slide1.xml').toString();
  assert.deepEqual(Array.from(xml.matchAll(/<a:bodyPr\b[^>]*\bvert="([^"]+)"/g), m => m[1]), ['vert', 'vert270', 'horz']);
  const imported = await importPptx(bytes, { projection: 'content' });
  try {
    const elements = imported.document.slides[0].elements;
    assert.deepEqual(elements.map(e => (e.type === 'shape' ? e.text.frame : e.frame).textDirection), ['rotate90', 'rotate270', 'horizontal']);
    const output = await renderPresentationToSvg(imported.document);
    assert.match(output.slides[0].svg, /rotate\(90 60 105\)/);
    assert.match(output.slides[0].svg, /rotate\(270 170 105\)/);
    assert.equal(output.warnings.filter(w => w.code === 'text-frame-direction-approximation').length, 2);
  } finally { imported.dispose(); }
});

test('three source text edits keep rotated body properties and neighboring parts intact', async () => {
  const deck = createPresentation();
  const slide = deck.addSlide();
  slide.addElement({ type: 'text', id: 'target', box: { x: 20, y: 30, width: 80, height: 150 }, frame: { textDirection: 'rotate90' }, content: '日本語 ABC 123' });
  deck.addSlide().addElement({ type: 'text', id: 'control', box: { x: 20, y: 30, width: 80, height: 150 }, frame: { textDirection: 'rotate270' }, content: '日本語 ABC 789' });
  let bytes = (await generatePptx(deck)).bytes;
  const part = 'ppt/slides/slide1.xml';
  for (const [oldDigits, newDigits] of [['123', '456'], ['456', '567'], ['567', '678']]) {
    const session = await importPptx(bytes, { projection: 'content', editSaveProfile: 'text-runs-v3' });
    try {
      const page = session.document.slides[0], before = page.elements[0];
      assert.equal(before.frame.textDirection, 'rotate90');
      assert.equal(before.content[0].runs[0].text, '日本語 ABC ' + oldDigits);
      const styleEdit = structuredClone(before);
      styleEdit.frame.textDirection = 'rotate270';
      page.replaceElement(before.id, styleEdit);
      assert.equal((await session.prepareSave({ mode: 'edit' })).status, 'blocked');
      const changed = structuredClone(before);
      changed.content[0].runs[0].text = '日本語 ABC ' + newDigits;
      page.replaceElement(before.id, changed);
      const prepared = await session.prepareSave({ mode: 'edit' });
      assert.equal(prepared.status, 'ready');
      const inputParts = readGeneratedZip(bytes);
      bytes = (await session.save(prepared.plan)).bytes;
      const outputParts = readGeneratedZip(bytes);
      assert.deepEqual([...outputParts.keys()].sort(), [...inputParts.keys()].sort());
      for (const [name, input] of inputParts) {
        if (name !== part) assert.deepEqual(outputParts.get(name), input, name);
      }
      const inputXml = inputParts.get(part).toString();
      assert.equal(inputXml.split(oldDigits).length, 2);
      assert.equal(outputParts.get(part).toString(), inputXml.replace(oldDigits, newDigits));
    } finally { session.dispose(); }
  }
  const final = await importPptx(bytes, { projection: 'content' });
  try {
    const element = final.document.slides[0].elements[0];
    assert.equal(element.frame.textDirection, 'rotate90');
    assert.equal(element.content[0].runs[0].text, '日本語 ABC 678');
  } finally { final.dispose(); }
});

test('East Asian vertical frame and cell preserve native properties without horizontal measurement', async () => {
  const deck = createPresentation(), slide = deck.addSlide();
  const box = { x: 20, y: 30, width: 100, height: 150 };
  const content = '日本語（ABC 123）', textDirection = 'eastAsianVertical';
  slide.addElement({ type: 'text', id: 'frame', box, content, frame: { textDirection } });
  slide.addElement({ type: 'table', id: 'table', box: { ...box, x: 150 }, columns: [100], rows: [
    { height: 150, cells: [{ content, style: { textDirection } }] },
  ] });
  assert.deepEqual(validatePresentation(deck), []);
  const result = resolveLayout(deck, { textMeasurement: { measure() { assert.fail('Horizontal measurement is not valid for vertical writing'); } } });
  assert.equal(result.analysis.measurements.length, 0);
  assert.equal(result.analysis.diagnostics.filter(d => d.code === 'east-asian-vertical-measurement-unavailable').length, 2);
  const sized = resolveLayout(deck, { textMeasurement: { measure() { assert.fail('A vertical cell must not be sized with horizontal metrics'); } },
    tableSizing: [{ elementId: 'table', autoRows: [{ row: 0 }] }] });
  assert.equal(sized.slides[0].elements[1].rows[0].height, 150);
  assert.ok(sized.analysis.diagnostics.some(d => d.code === 'east-asian-vertical-measurement-unavailable' && d.elementId === 'table'));
  const output = await generatePptx(deck), xml = readGeneratedZip(output.bytes).get('ppt/slides/slide1.xml').toString();
  assert.match(xml, /<a:bodyPr\b[^>]* vert="eaVert"/);
  assert.match(xml, /<a:tcPr\b[^>]* vert="eaVert"/);
  const session = await importPptx(output.bytes, { projection: 'content' });
  try {
    const [frame, table] = session.document.slides[0].elements;
    assert.equal(frame.frame.textDirection, textDirection);
    assert.equal(table.rows[0].cells[0].style.textDirection, textDirection);
    const preview = await renderPresentationToSvg(session.document);
    assert.equal(preview.slides[0].svg.match(/writing-mode:vertical-rl;text-orientation:mixed/g)?.length, 2);
    assert.equal(preview.warnings.filter(w => w.code === 'east-asian-vertical-preview-approximation').length, 2);
  } finally { session.dispose(); }
});

test('East Asian vertical frame and cell edits preserve both property owners and the horizontal control', async () => {
  const deck = createPresentation(), page = deck.addSlide();
  const textDirection = 'eastAsianVertical';
  page.addElement({ type: 'text', id: 'v', box: { x: 20, y: 30, width: 120, height: 200 },
    frame: { textDirection }, content: '日本語（ABC 123）' });
  page.addElement({ type: 'table', id: 't', box: { x: 160, y: 30, width: 240, height: 200 }, columns: [120, 120], rows: [
    { height: 200, cells: [{ content: '日本語（DEF 234）', style: { textDirection } }, { content: '対照（GHI 345）' }] },
  ] });
  const original = (await generatePptx(deck)).bytes, part = 'ppt/slides/slide1.xml';
  const session = await importPptx(original, { projection: 'content', editSaveProfile: 'text-runs-v3' });
  let saved;
  try {
    const slide = session.document.slides[0];
    const [frame, table] = slide.elements.map(e => structuredClone(e));
    frame.content[0].runs[0].text = '日本語（ABC 456）';
    table.rows[0].cells[0].content[0].runs[0].text = '日本語（DEF 567）';
    slide.replaceElement(frame.id, frame);
    slide.replaceElement(table.id, table);
    const ready = await session.prepareSave({ mode: 'edit' });
    assert.equal(ready.status, 'ready');
    saved = (await session.save(ready.plan)).bytes;
    const input = readGeneratedZip(original), output = readGeneratedZip(saved);
    assert.deepEqual([...output.keys()].sort(), [...input.keys()].sort());
    for (const [name, bytes] of input) if (name !== part) assert.deepEqual(output.get(name), bytes, name);
    assert.equal(output.get(part).toString(), input.get(part).toString().replace('ABC 123', 'ABC 456').replace('DEF 234', 'DEF 567'));
  } finally { session.dispose(); }
  const reopened = await importPptx(saved, { projection: 'content' });
  try {
    const [frame, table] = reopened.document.slides[0].elements;
    assert.equal(frame.frame.textDirection, textDirection);
    assert.equal(frame.content[0].runs[0].text, '日本語（ABC 456）');
    const [vertical, control] = table.rows[0].cells;
    assert.equal(vertical.style.textDirection, textDirection);
    assert.equal(vertical.content[0].runs[0].text, '日本語（DEF 567）');
    assert.equal(control.style.textDirection, undefined);
    assert.equal(control.content[0].runs[0].text, '対照（GHI 345）');
  } finally { reopened.dispose(); }
});

for (const [direction, native, writingMode, orientation] of [
  ['mongolianVertical', 'mongolianVert', 'vertical-lr', 'mixed'],
  ['stackedVertical', 'wordArtVert', 'vertical-lr', 'upright'],
  ['stackedVerticalRtl', 'wordArtVertRtl', 'vertical-rl', 'upright'],
]) test(`${native} retains distinct frame and cell writing modes`, async () => {
  const deck = createPresentation(), slide = deck.addSlide();
  const content = '日本語 ABC 123', box = { x: 20, y: 30, width: 120, height: 200 };
  slide.addElement({ type: 'text', id: 'v', box, content, frame: { textDirection: direction } });
  slide.addElement({ type: 'table', id: 't', box: { ...box, x: 160 }, columns: [120], rows: [
    { height: 200, cells: [{ content, style: { textDirection: direction } }] },
  ] });
  const bytes = (await generatePptx(deck)).bytes;
  const xml = readGeneratedZip(bytes).get('ppt/slides/slide1.xml').toString();
  for (const tag of ['bodyPr', 'tcPr']) assert.match(xml, new RegExp(`<a:${tag}[^>]* vert="${native}"`));
  const imported = await importPptx(bytes, { projection: 'content' });
  try {
    const [frame, table] = imported.document.slides[0].elements;
    assert.equal(frame.frame.textDirection, direction);
    assert.equal(table.rows[0].cells[0].style.textDirection, direction);
    const preview = await renderPresentationToSvg(imported.document);
    assert.equal(preview.slides[0].svg.split(`writing-mode:${writingMode};text-orientation:${orientation}`).length, 3);
    assert.equal(preview.warnings.filter(w => w.code === 'vertical-text-preview-approximation').length, 2);
    const analysis = resolveLayout(imported.document, { textMeasurement: { measure() { assert.fail('Horizontal measurement is not valid'); } },
      tableSizing: [{ elementId: table.id, autoRows: [{ row: 0 }] }] }).analysis;
    assert.equal(analysis.measurements.length, 0);
    assert.ok(analysis.diagnostics.some(d => d.code === 'vertical-text-measurement-unavailable' && d.elementId === frame.id));
    assert.ok(analysis.diagnostics.some(d => d.code === 'vertical-text-measurement-unavailable' && d.elementId === table.id));
  } finally { imported.dispose(); }
});

test('native frame direction inherits from the layout and retains local horizontal overrides', async () => {
  const deck = createPresentation({ textStylePresets: { horizontal: { frame: { textDirection: 'horizontal' } } } });
  deck.defineSlideLayout({ id: 'layout', name: 'Vertical layout', placeholders: [{
    key: 'body', kind: 'body', box: { x: 20, y: 30, width: 120, height: 200 }, textStyle: { frame: { textDirection: 'eastAsianVertical' } },
  }] });
  for (const local of [false, true]) {
    const slide = deck.addSlide({ layoutId: 'layout' });
    slide.addElement({ type: 'text', id: local ? 'local' : 'inherited', placeholderKey: 'body', content: '中文 ABC 123',
      textStylePreset: 'horizontal', textStylePresetReset: ['frame.textDirection'],
      ...(local ? { frame: { textDirection: 'horizontal' } } : {}) });
  }
  const output = await generatePptx(deck, { nativePlaceholderText: 'inherit' }), parts = readGeneratedZip(output.bytes);
  assert.doesNotMatch(parts.get('ppt/slides/slide1.xml').toString(), /<a:bodyPr[^>]* vert=/);
  assert.match(parts.get('ppt/slides/slide2.xml').toString(), /<a:bodyPr[^>]* vert="horz"/);
  assert.ok([...parts].some(([name, bytes]) => name.startsWith('ppt/slideLayouts/slideLayout') && /<a:bodyPr[^>]* vert="eaVert"/.test(bytes.toString())));
  assert.ok(!output.warnings.some(w => w.code === 'native-text-inheritance-materialized'));
  const session = await importPptx(output.bytes, { projection: 'content' });
  try {
    assert.deepEqual(session.document.slides.map(s => s.elements[0].frame.textDirection), ['eastAsianVertical', 'horizontal']);
    const directions = session.report.textStyleInheritance.properties.filter(p => p.property === 'frame.textDirection');
    assert.ok(directions.some(p => p.level === 'layout-placeholder' && p.value === 'eastAsianVertical'));
    assert.ok(directions.some(p => p.level !== 'layout-placeholder' && p.value === 'horizontal'));
  } finally { session.dispose(); }
});

test('Mongolian variation and separation controls survive script-font materialization', async () => {
  const deck = createPresentation();
  const mongolian = 'ᠠ\u180B ᠠ\u180C ᠠ\u180D ᠠ\u180Eᠠ';
  const text = 'A ' + mongolian + ' Z';
  deck.addSlide().addElement({ type: 'text', id: 'variants', box: { x: 0, y: 0, width: 120, height: 300 },
    frame: { textDirection: 'mongolianVertical' }, content: [{ runs: [{ text, style: {
      fontFamily: 'Arial', scriptFontFamilies: { Mong: 'Noto Sans Mongolian' }, language: 'mn-Mong',
    } }] }] });
  const exported = await generatePptx(deck);
  const xml = readGeneratedZip(exported.bytes).get('ppt/slides/slide1.xml').toString();
  assert.match(xml, /<a:bodyPr[^>]* vert="mongolianVert"/);
  assert.deepEqual([...xml.matchAll(/<a:t>([^<]*)<\/a:t>/gu)].map(m => m[1]), ['A ', mongolian + ' ', 'Z']);
  const session = await importPptx(exported.bytes, { projection: 'content', textFontProfile: 'script-families-v1' });
  try {
    const frame = session.document.slides[0].elements[0], runs = frame.content[0].runs;
    assert.equal(runs.map(run => run.text).join(''), text);
    assert.equal(runs[1].style.fontFamily, 'Noto Sans Mongolian');
    assert.equal(frame.frame.textDirection, 'mongolianVertical');
    const result = await renderPresentationToSvg(deck);
    assert.ok(result.slides[0].svg.includes(mongolian + ' '));
  } finally { session.dispose(); }
});
