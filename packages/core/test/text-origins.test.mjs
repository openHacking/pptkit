import assert from 'node:assert/strict';
import test from 'node:test';
import { createPresentation, normalizePresentation, resolvePresentationInheritance } from '../dist/index.js';
import { textOriginsDocument } from '../../../tests/fixtures/text-origins.mjs';

function row(r, id, property) {
  const matches = r.origins.filter(o => o.target.kind === 'element' && o.target.id === id && o.property === property);
  assert.equal(matches.length, 1, `${id}/${property}`); return matches[0];
}
const builtin = path => ({ kind: 'default', path });
const preset = (id, suffix) => ({ kind: 'preset', id, path: `textStylePresets[${JSON.stringify(id)}].${suffix}` });
const placeholder = suffix => ({ kind: 'placeholder', id: 'a', key: 'same', path: `layouts.0.placeholders.0.${suffix}` });
const local = suffix => ({ kind: 'element', id: 'body', path: `slides.0.elements.0.${suffix}` });

test('text origins retain the actual placeholder/preset/local cascade alongside identical ordinary IR', () => {
  const d = textOriginsDocument(createPresentation), r = resolvePresentationInheritance(d);
  assert.deepEqual(r.presentation, normalizePresentation(d));
  assert.deepEqual(row(r, 'body', 'content.0.runs.1.style.fontSize').source, preset('base.with.dot', 'run.fontSize'));
  assert.deepEqual(row(r, 'body', 'content.0.runs.1.style.italic').source, preset('element"preset', 'run.italic'));
  assert.deepEqual(row(r, 'body', 'content.0.runs.1.style.color').source, placeholder('textStyle.run.color'));
  assert.deepEqual(row(r, 'body', 'content.0.runs.0.style.bold').source, local('content.0.runs.0.style.bold'));
  const body = r.presentation.slides[0].elements[0];
  assert.equal(body.content[0].runs[0].text, ''); assert.equal(body.content[0].runs[0].style.bold, false); assert.equal(body.content[0].runs[0].style.colorOpacity, 0);
  assert.deepEqual(row(r, 'body', 'content.0.runs.0.style.colorOpacity').source, local('content.0.runs.0.style.colorOpacity'));
});

test('placeholder targets distinguish identical keys across layouts and shared layout elements are recorded once', () => {
  const r = resolvePresentationInheritance(textOriginsDocument(createPresentation));
  const rows = r.origins.filter(o => o.target.kind === 'placeholder' && o.property === 'textStyle.run.color');
  assert.deepEqual(rows.map(o => o.target), [{ kind: 'placeholder', id: 'a', key: 'same' }, { kind: 'placeholder', id: 'b', key: 'same' }]);
  assert.deepEqual(row(r, 'other', 'content.0.runs.0.style.color').source, { kind: 'placeholder', id: 'b', key: 'same', path: 'layouts.1.placeholders.0.textStyle.run.color' });
  assert.deepEqual(row(r, 'shared-text', 'content.0.runs.0.style.color').source, placeholder('textStyle.run.color'));
  const identities = r.origins.map(o => JSON.stringify([o.target, o.property])); assert.equal(new Set(identities).size, identities.length);
});

test('frame scalar and partial margins, explicit false and whole autoFit replacement retain exact sources', () => {
  const r = resolvePresentationInheritance(textOriginsDocument(createPresentation)), frame = r.presentation.slides[0].elements[0].frame;
  assert.deepEqual(frame.margin, { top: 0, right: 0, bottom: 0, left: 0 }); assert.equal(frame.wrap, false); assert.deepEqual(frame.autoFit, { mode: 'shrink', fontScale: 1 });
  for (const side of ['top', 'right', 'bottom', 'left']) assert.deepEqual(row(r, 'body', `frame.margin.${side}`).source, local('frame.margin'));
  assert.deepEqual(row(r, 'shape', 'text.frame.margin.left').source, placeholder('textStyle.frame.margin.left'));
  assert.deepEqual(row(r, 'shape', 'text.frame.margin.right').source, preset('base.with.dot', 'frame.margin'));
  assert.deepEqual(row(r, 'body', 'frame.wrap').source, local('frame.wrap'));
  assert.deepEqual(row(r, 'body', 'frame.autoFit.fontScale').source, builtin('defaults.text.frame.autoFit.fontScale'));
});

test('list synthesis reports built-in indentation and independent explicit zero rather than inherited spacing', () => {
  const r = resolvePresentationInheritance(textOriginsDocument(createPresentation)), paragraph = r.presentation.slides[0].elements[0].content[0].style;
  assert.equal(paragraph.indent, 0); assert.equal(paragraph.hanging, 27); assert.equal(paragraph.spaceBefore, 0);
  assert.deepEqual(row(r, 'body', 'content.0.style.indent').source, local('content.0.style.indent'));
  assert.deepEqual(row(r, 'body', 'content.0.style.hanging').source, builtin('defaults.text.paragraph.listIndent'));
  assert.deepEqual(row(r, 'body', 'content.0.style.bullet.startAt').source, builtin('defaults.text.paragraph.bullet.startAt'));
  assert.deepEqual(row(r, 'body', 'content.0.style.bullet.type').source, local('content.0.style.bullet.type'));
  assert.equal(r.origins.some(o => o.target.kind === 'element' && o.target.id === 'shape' && o.property === 'text.content.0.style.bullet.startAt'), false);
});

test('whole-map and per-script clears survive fallback normalization while a later slot can reopen', () => {
  const r = resolvePresentationInheritance(textOriginsDocument(createPresentation)), body = r.presentation.slides[0].elements[0];
  assert.equal(body.content[0].runs[0].style.scriptFontFamilies, undefined);
  assert.deepEqual(body.content[0].runs[1].style.scriptFontFamilies, { Hans: 'Verdana' });
  assert.equal(row(r, 'body', 'content.0.runs.0.style.scriptFontFamilies.Hans').state, 'cleared');
  assert.deepEqual(row(r, 'body', 'content.0.runs.0.style.scriptFontFamilies.Hans').source, local('content.0.runs.0.style.scriptFontFamilies.Hans'));
  assert.deepEqual(row(r, 'body', 'content.0.runs.1.style.scriptFontFamilies.Hans').source, preset('element"preset', 'run.scriptFontFamilies.Hans'));
  assert.equal(row(r, 'body', 'content.0.runs.1.style.scriptFontFamilies.Arab').state, 'cleared');
  assert.deepEqual(row(r, 'body', 'content.0.runs.1.style.scriptFontFamilies.Arab').source, placeholder('textStyle.run.scriptFontFamilies'));
  assert.equal(row(r, 'shape', 'text.content.0.runs.0.style.scriptFontFamilies.Hans').state, 'cleared');
  assert.equal(row(r, 'body', 'content.0.runs.1.style.scriptFontFamilies.Hans').state, undefined);
});

test('optional font clears are distinguishable from never-declared fields, including on empty text', () => {
  const r = resolvePresentationInheritance(textOriginsDocument(createPresentation));
  assert.equal(row(r, 'body', 'content.0.runs.1.style.eastAsianFontFamily').state, 'cleared');
  assert.deepEqual(row(r, 'body', 'content.0.runs.1.style.eastAsianFontFamily').source, placeholder('textStyle.run.eastAsianFontFamily'));
  assert.equal(row(r, 'body', 'content.0.runs.0.style.complexScriptFontFamily').state, 'cleared');
  assert.deepEqual(row(r, 'body', 'content.0.runs.0.style.complexScriptFontFamily').source, local('content.0.runs.0.style.complexScriptFontFamily'));
  assert.equal(r.origins.some(o => o.target.kind === 'element' && o.target.id === 'child' && o.property.endsWith('eastAsianFontFamily')), false);
  assert.deepEqual(row(r, 'child', 'content.0.runs.0.style.bold').source, builtin('defaults.text.run.bold'));
});

test('group children stay outside the parent placeholder cascade and shape text retains its own paths', () => {
  const r = resolvePresentationInheritance(textOriginsDocument(createPresentation));
  assert.deepEqual(row(r, 'child', 'frame.margin.left').source, builtin('defaults.text.frame.margin.left'));
  assert.deepEqual(row(r, 'child', 'content.0.runs.0.style.color').source, builtin('defaults.text.run.color'));
  assert.deepEqual(row(r, 'child', 'frame.wrap').source, { kind: 'element', id: 'child', path: 'slides.0.elements.2.children.0.frame.wrap' });
  assert.deepEqual(row(r, 'shape', 'text.content.0.runs.0.style.fontSize').source, preset('base.with.dot', 'run.fontSize'));
  assert.equal(r.origins.some(o => o.target.kind === 'element' && o.target.id === 'group'), false);
});

test('equal false/default values retain explicit declarations and reports remain detached immutable snapshots', () => {
  const d = textOriginsDocument(createPresentation), before = resolvePresentationInheritance(d), saved = JSON.stringify(before.origins);
  assert.throws(() => { row(before, 'body', 'content.0.runs.0.style.bold').source.path = 'bad'; }, TypeError);
  const slide = d.slides[0], body = slide.elements[0];
  slide.replaceElement('body', { ...body, content: [{ runs: [{ text: '' }] }] });
  const after = resolvePresentationInheritance(d);
  assert.equal(after.presentation.slides[0].elements[0].content[0].runs[0].style.bold, true);
  assert.deepEqual(row(after, 'body', 'content.0.runs.0.style.bold').source, preset('base.with.dot', 'run.bold'));
  assert.equal(JSON.stringify(before.origins), saved);
  const plain = createPresentation(); plain.addSlide({ id: 's', elements: [{ id: 'equal', type: 'text', box: { x: 0, y: 0, width: 100, height: 20 }, content: [{ runs: [{ text: '', style: { bold: false, colorOpacity: 1 } }] }] }] });
  const equal = resolvePresentationInheritance(plain);
  assert.equal(equal.presentation.slides[0].elements[0].content[0].runs[0].style.bold, false);
  assert.equal(equal.presentation.slides[0].elements[0].content[0].runs[0].style.colorOpacity, undefined);
  for (const property of ['bold', 'colorOpacity']) assert.deepEqual(row(equal, 'equal', `content.0.runs.0.style.${property}`).source, { kind: 'element', id: 'equal', path: `slides.0.elements.0.content.0.runs.0.style.${property}` });
});

test('every recorded authored formatting source reads the actual selected leaf or explicit null barrier', () => {
  const d = textOriginsDocument(createPresentation), r = resolvePresentationInheritance(d), elements = new Map();
  const collect = values => { for (const e of values) { elements.set(e.id, e); if (e.type === 'group') collect(e.children); } };
  for (const layout of r.presentation.layouts) collect(layout.elements);
  for (const slide of r.presentation.slides) collect(slide.elements);
  const leaves = (value, prefix) => {
    const atomic = /\.(?:color|fontFamily|eastAsianFontFamily|complexScriptFontFamily)$/.test(prefix) || /\.scriptFontFamilies\.[^.]+$/.test(prefix);
    return value !== null && typeof value === 'object' && !atomic ? Object.entries(value).flatMap(([key, child]) => leaves(child, `${prefix}.${key}`)) : [prefix];
  };
  const coverage = (target, value, prefix, run = false) => {
    const required = leaves(value, prefix); if (run && value.colorOpacity === undefined) required.push(`${prefix}.colorOpacity`);
    for (const property of required) assert.equal(r.origins.filter(o => o.target.kind === target.kind && o.target.id === target.id && o.target.key === target.key && o.property === property).length, 1, `Missing effective property ${JSON.stringify(target)}/${property}`);
  };
  for (const e of elements.values()) {
    const target = { kind: 'element', id: e.id }, text = e.type === 'text' ? e : e.type === 'shape' ? e.text : undefined, prefix = e.type === 'shape' ? 'text.' : '';
    if (!text) continue;
    coverage(target, text.frame, `${prefix}frame`);
    text.content.forEach((p, pi) => { coverage(target, p.style, `${prefix}content.${pi}.style`); p.runs.forEach((run, ri) => coverage(target, run.style, `${prefix}content.${pi}.runs.${ri}.style`, true)); });
  }
  for (const l of r.presentation.layouts) for (const p of l.placeholders) for (const component of ['frame', 'paragraph', 'run']) coverage({ kind: 'placeholder', id: l.id, key: p.key }, p.textStyle[component], `textStyle.${component}`, component === 'run');
  const read = (value, path) => path.split('.').reduce((current, key) => current?.[key], value);
  const sourceValue = s => s.kind === 'preset' ? read(d.textStylePresets[s.id], s.path.slice(`textStylePresets[${JSON.stringify(s.id)}].`.length)) : read(d, s.path);
  let checked = 0, clears = 0;
  for (const o of r.origins) {
    if (!['element', 'placeholder'].includes(o.target.kind) || o.source.kind === 'default') continue;
    const declared = sourceValue(o.source);
    if (o.state === 'cleared') { assert.equal(declared, null, JSON.stringify(o)); clears++; continue; }
    const target = o.target.kind === 'element' ? elements.get(o.target.id) : r.presentation.layouts.find(l => l.id === o.target.id).placeholders.find(p => p.key === o.target.key);
    const effective = read(target, o.property);
    assert.deepEqual(declared, effective === undefined && o.property.endsWith('.colorOpacity') ? 1 : effective, JSON.stringify(o));
    checked++;
  }
  assert.ok(checked > 30); assert.ok(clears > 30);
});

test('removing shared placeholder declarations restores defaults and removes old clear barriers on re-resolution', () => {
  const d = textOriginsDocument(createPresentation), before = resolvePresentationInheritance(d), old = d.layouts[0];
  d.replaceSlideLayout('a', { name: 'Reset', elements: [...old.elements], placeholders: old.placeholders.map(p => ({ key: p.key, kind: p.kind, box: p.box })) });
  const after = resolvePresentationInheritance(d), body = after.presentation.slides[0].elements[0];
  assert.equal(body.content[0].runs[1].style.fontSize, 18);
  assert.deepEqual(row(after, 'body', 'content.0.runs.1.style.fontSize').source, builtin('defaults.text.run.fontSize'));
  assert.equal(after.origins.some(o => o.target.kind === 'element' && o.target.id === 'body' && o.property === 'content.0.runs.1.style.eastAsianFontFamily'), false);
  assert.equal(row(before, 'body', 'content.0.runs.1.style.eastAsianFontFamily').state, 'cleared');
});
