import assert from 'node:assert/strict';
import test from 'node:test';
import { createPresentation, normalizePresentation, resolvePresentationInheritance } from '../dist/index.js';
import { chartNotesOriginsDocument } from '../../../tests/fixtures/chart-notes-origins.mjs';
const fixture = () => chartNotesOriginsDocument(createPresentation);
function origin(r, kind, id, property) {
  const rows = r.origins.filter(o => o.target.kind === kind && o.target.id === id && o.property === property);
  assert.equal(rows.length, 1, `${kind}/${id}/${property}`); return rows[0];
}
const local = (id, path) => ({ kind: 'element', id, path });
const builtin = path => ({ kind: 'default', path });

test('chart style origins retain source declarations while resolving each owning master theme', () => {
  const d = fixture(), r = resolvePresentationInheritance(d);
  assert.deepEqual(r.presentation, normalizePresentation(d));
  for (const [index, fontFamily, color] of [[0, 'Georgia', '#CC2200'], [1, 'Verdana', '#0044CC'], [2, 'Georgia', '#CC2200']]) for (const [ci, type] of ['bar', 'line', 'pie'].entries()) {
    const e = r.presentation.slides[index].elements[ci], id = type + index, path = `slides.${index}.elements.${ci}`;
    assert.deepEqual(e.style.textStyle, { fontFamily, fontSize: 14, bold: true, italic: true, color });
    assert.deepEqual(origin(r, 'element', id, 'style.textStyle.fontFamily').source, local(id, path + '.style.textStyle.fontFamily'));
    assert.deepEqual(origin(r, 'element', id, 'legend.textStyle.color').source, local(id, path + '.style.textStyle.color'));
    assert.deepEqual(origin(r, 'element', id, 'legend.textStyle.italic').source, local(id, path + '.legend.textStyle.italic'));
    assert.equal(e.legend.textStyle.italic, false); assert.equal(e.legend.visible, false);
    assert.deepEqual(e.titleStyle, { fontFamily, color: '#112233', fontSize: 12, bold: false, italic: true });
    assert.deepEqual(origin(r, 'element', id, 'titleStyle.fontSize').source, builtin('defaults.chart.titleStyle.fontSize'));
    assert.deepEqual(origin(r, 'element', id, 'titleStyle.bold').source, local(id, path + '.titleStyle.bold'));
    if (type !== 'pie') {
      assert.equal(e.axes.category.labelStyle.fontSize, 9); assert.equal(e.axes.value.labelStyle.fontFamily, 'Arial');
      assert.deepEqual(origin(r, 'element', id, 'axes.category.labelStyle.fontSize').source, local(id, path + '.axes.category.labelStyle.fontSize'));
      assert.deepEqual(origin(r, 'element', id, 'axes.value.labelStyle.italic').source, local(id, path + '.style.textStyle.italic'));
      assert.deepEqual(origin(r, 'element', id, 'axes.value.labelStyle.color').source, local(id, path + '.axes.value.labelStyle.color'));
    } else assert.equal(r.origins.some(o => o.target.kind === 'element' && o.target.id === id && o.property.startsWith('axes.')), false);
  }
});

test('chart defaults and title-specific defaults stay distinct, including absent titles and explicit equal values', () => {
  const r = resolvePresentationInheritance(fixture()), e = r.presentation.slides[3].elements[0];
  assert.equal(e.title, undefined); assert.equal(e.titleStyle.bold, true); assert.equal(e.titleStyle.fontSize, 12);
  for (const prefix of ['style.textStyle', 'legend.textStyle']) for (const key of ['fontFamily', 'fontSize', 'bold', 'italic', 'color']) assert.deepEqual(origin(r, 'element', e.id, `${prefix}.${key}`).source, builtin(`defaults.chart.textStyle.${key}`));
  assert.deepEqual(origin(r, 'element', e.id, 'titleStyle.bold').source, builtin('defaults.chart.titleStyle.bold'));
  assert.deepEqual(origin(r, 'element', 'nested-chart', 'titleStyle.fontSize').source, local('nested-chart', 'slides.3.elements.1.children.0.titleStyle.fontSize'));
  assert.deepEqual(origin(r, 'element', 'nested-chart', 'titleStyle.bold').source, local('nested-chart', 'slides.3.elements.1.children.0.titleStyle.bold'));
});

test('notes have slide-scoped targets, direct source paths and no placeholder/preset/chart inheritance', () => {
  const r = resolvePresentationInheritance(fixture());
  for (let index = 0; index < 3; index++) {
    const id = 'page-' + index, notes = r.presentation.slides[index].notes, owner = suffix => ({ kind: 'slide', id, path: `slides.${index}.notes.${suffix}` });
    assert.equal(notes[0].runs[0].style.fontSize, 18); assert.equal(notes[0].runs[0].style.bold, false); assert.equal(notes[0].runs[0].style.colorOpacity, 0);
    assert.deepEqual(origin(r, 'notes', id, 'notes.0.runs.0.style.fontSize').source, builtin('defaults.text.run.fontSize'));
    assert.deepEqual(origin(r, 'notes', id, 'notes.0.runs.0.style.bold').source, owner('0.runs.0.style.bold'));
    for (const property of ['eastAsianFontFamily', 'scriptFontFamilies.Hans', 'scriptFontFamilies.Arab']) assert.equal(origin(r, 'notes', id, `notes.0.runs.0.style.${property}`).state, 'cleared');
    assert.deepEqual(origin(r, 'notes', id, 'notes.0.runs.0.style.scriptFontFamilies.Hans').source, owner('0.runs.0.style.scriptFontFamilies'));
    assert.deepEqual(origin(r, 'notes', id, 'notes.0.runs.1.style.scriptFontFamilies.Hans').source, owner('0.runs.1.style.scriptFontFamilies.Hans'));
    assert.equal(origin(r, 'notes', id, 'notes.0.runs.1.style.scriptFontFamilies.Arab').state, 'cleared');
    assert.deepEqual(origin(r, 'notes', id, 'notes.0.style.indent').source, owner('0.style.indent'));
    assert.deepEqual(origin(r, 'notes', id, 'notes.0.style.hanging').source, builtin('defaults.text.paragraph.listIndent'));
    assert.equal(r.origins.some(o => o.target.kind === 'notes' && o.target.id === id && ['preset', 'placeholder', 'master', 'element'].includes(o.source.kind)), false);
  }
});

test('missing notes, empty strings/paragraphs/runs and source positions retain actual normalization semantics', () => {
  const d = fixture(), slide = d.slides[3], r = resolvePresentationInheritance(d);
  assert.equal(r.presentation.slides[3].notes[0].runs[0].text, '');
  assert.deepEqual(origin(r, 'notes', slide.id, 'notes.0.runs.0.style.bold').source, builtin('defaults.text.run.bold'));
  slide.setNotes('first\r\n\rthird'); const multiline = resolvePresentationInheritance(d);
  assert.equal(multiline.presentation.slides[3].notes.length, 3);
  for (let p = 0; p < 3; p++) assert.deepEqual(origin(multiline, 'notes', slide.id, `notes.${p}.runs.0.style.fontFamily`).source, builtin('defaults.text.run.fontFamily'));
  slide.setNotes([{ runs: [] }, { runs: [{ text: '', style: { language: 'zh-CN', colorOpacity: 1 } }] }]);
  const empty = resolvePresentationInheritance(d);
  assert.equal(empty.origins.some(o => o.target.kind === 'notes' && o.target.id === slide.id && o.property.startsWith('notes.0.runs.')), false);
  assert.deepEqual(origin(empty, 'notes', slide.id, 'notes.1.runs.0.style.colorOpacity').source, { kind: 'slide', id: slide.id, path: 'slides.3.notes.1.runs.0.style.colorOpacity' });
  slide.setNotes([]); const none = resolvePresentationInheritance(d);
  assert.equal(none.origins.some(o => o.target.kind === 'notes' && o.target.id === slide.id), false);
});

test('notes/chart re-resolution clears stale origins after reset and keeps prior snapshots immutable', () => {
  const d = fixture(), before = resolvePresentationInheritance(d), saved = JSON.stringify(before.origins), slide = d.slides[0];
  slide.setNotes();
  const e = slide.elements[0]; slide.replaceElement(e.id, { ...e, style: undefined, titleStyle: undefined, legend: undefined, axes: undefined });
  d.replaceSlideMaster('red', { name: 'New red', theme: { fonts: { body: 'Tahoma' }, colors: { accent1: '00FF00' } } });
  const after = resolvePresentationInheritance(d);
  assert.equal(after.presentation.slides[0].elements[0].style.textStyle.fontFamily, 'Tahoma');
  assert.deepEqual(origin(after, 'element', 'bar0', 'style.textStyle.fontFamily').source, builtin('defaults.chart.textStyle.fontFamily'));
  assert.equal(after.origins.some(o => o.target.kind === 'notes' && o.target.id === slide.id && o.state === 'cleared'), false);
  assert.equal(JSON.stringify(before.origins), saved);
  assert.throws(() => { origin(before, 'notes', slide.id, 'notes.0.runs.0.style.bold').source.path = 'bad'; }, TypeError);
  assert.deepEqual(after.presentation, normalizePresentation(d));
});

test('every supported chart and notes formatting leaf has exactly one source, without duplicate targets', () => {
  const r = resolvePresentationInheritance(fixture());
  const ids = r.origins.map(o => JSON.stringify([o.target, o.property])); assert.equal(new Set(ids).size, ids.length);
  for (const slide of r.presentation.slides) {
    slide.notes.forEach((p, pi) => {
      for (const key of Object.entries(p.style).flatMap(([k, v]) => typeof v === 'object' ? Object.keys(v).map(child => `${k}.${child}`) : [k])) origin(r, 'notes', slide.id, `notes.${pi}.style.${key}`);
      p.runs.forEach((run, ri) => { for (const key of [...Object.keys(run.style), ...(run.style.colorOpacity === undefined ? ['colorOpacity'] : [])]) {
        if (key === 'scriptFontFamilies') for (const code of Object.keys(run.style.scriptFontFamilies)) origin(r, 'notes', slide.id, `notes.${pi}.runs.${ri}.style.scriptFontFamilies.${code}`);
        else origin(r, 'notes', slide.id, `notes.${pi}.runs.${ri}.style.${key}`);
      } });
    });
    const walk = elements => { for (const e of elements) { if (e.type === 'group') walk(e.children); if (e.type !== 'chart') continue;
      for (const [prefix, value] of [['style.textStyle', e.style.textStyle], ['titleStyle', e.titleStyle], ['legend.textStyle', e.legend.textStyle], ...(e.chartType === 'pie' ? [] : [['axes.category.labelStyle', e.axes.category.labelStyle], ['axes.value.labelStyle', e.axes.value.labelStyle]])]) for (const key of Object.keys(value)) origin(r, 'element', e.id, `${prefix}.${key}`);
    } }; walk(slide.elements);
  }
});


test('authored source leaves independently explain chart theme references and notes null barriers', () => {
  const d = fixture(), r = resolvePresentationInheritance(d);
  const read = (root, path) => path.split('.').reduce((v, key) => v?.[key], root);
  const elements = new Map(); const walk = values => { for (const e of values) { elements.set(e.id, e); if (e.type === 'group') walk(e.children); } };
  r.presentation.slides.forEach(s => walk(s.elements));
  let symbolicFonts = 0, symbolicColors = 0, direct = 0, clears = 0;
  for (const o of r.origins) {
    if (!['notes', 'element'].includes(o.target.kind) || o.source.kind === 'default') continue;
    const target = o.target.kind === 'notes' ? r.presentation.slides.find(s => s.id === o.target.id) : elements.get(o.target.id);
    const declared = read(d, o.source.path), effective = read(target, o.property);
    if (o.state === 'cleared') { assert.equal(declared, null); assert.equal(effective, undefined); clears++; continue; }
    if (target.type === 'chart' && declared?.theme) {
      const slide = r.presentation.slides.find(s => s.elements.includes(target)), theme = r.presentation.layouts.find(l => l.id === slide.layoutId).theme;
      if (o.property.endsWith('.fontFamily')) { assert.equal(effective, theme.fonts[declared.theme]); symbolicFonts++; }
      else { assert.equal(effective, '#' + theme.colors[declared.theme]); symbolicColors++; }
    } else {
      const expected = target.type === 'chart' && o.property.endsWith('.color') && !declared.startsWith('#') ? '#' + declared : declared;
      assert.deepEqual(effective === undefined && o.property.endsWith('.colorOpacity') ? 1 : effective, expected, JSON.stringify(o)); direct++;
    }
  }
  assert.ok(symbolicFonts >= 30); assert.ok(symbolicColors >= 20); assert.ok(direct > 50); assert.ok(clears > 30);
});

test('layout chart sources are shared once and nested source paths never borrow slide chart formatting', () => {
  const d = fixture(), layout = d.layouts[0], source = d.slides[0].elements[0];
  d.replaceSlideLayout(layout.id, { ...layout, elements: [{ ...source, id: 'shared-chart' }] });
  const r = resolvePresentationInheritance(d);
  assert.deepEqual(origin(r, 'element', 'shared-chart', 'legend.textStyle.fontFamily').source, local('shared-chart', 'layouts.0.elements.0.style.textStyle.fontFamily'));
  assert.equal(r.presentation.layouts.find(l => l.id === layout.id).elements[0].style.textStyle.fontFamily, 'Georgia');
  const identities = r.origins.map(o => JSON.stringify([o.target, o.property])); assert.equal(new Set(identities).size, identities.length);
});
