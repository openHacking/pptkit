import assert from 'node:assert/strict';
import test from 'node:test';
import { createPresentation, normalizePresentation, resolvePresentationInheritance } from '../dist/index.js';
import { multiMasterDocument } from '../../../tests/fixtures/multi-master.mjs';

function source(result, kind, id, property) {
  const rows = result.origins.filter(r => r.target.kind === kind && r.target.id === id && r.property === property);
  assert.equal(rows.length, 1, `One origin for ${kind}/${id}/${property}`);
  return rows[0].source;
}
const declared = (kind, id, path) => ({ kind, id, path });
const builtin = path => ({ kind: 'default', path });

test('origin-enabled normalization returns identical ordinary IR without report fields', () => {
  const d = multiMasterDocument(createPresentation), r = resolvePresentationInheritance(d);
  assert.deepEqual(r.presentation, normalizePresentation(d));
  assert.equal(Object.hasOwn(r.presentation, 'origins'), false);
  assert.deepEqual(Object.keys(r), ['presentation', 'origins']);
  assert.ok(Object.isFrozen(r)); assert.ok(Object.isFrozen(r.origins));
  assert.ok(r.origins.every(row => Object.isFrozen(row) && Object.isFrozen(row.target) && Object.isFrozen(row.source)));
  assert.throws(() => { r.origins[0].source.path = 'bad'; }, TypeError);
  assert.throws(() => r.origins.pop(), TypeError);
});

test('theme roles track actual master/document/default declarations across interleaved domains', () => {
  const d = multiMasterDocument(createPresentation), r = resolvePresentationInheritance(d);
  assert.deepEqual(source(r, 'document', d.id, 'theme.colors.accent1'), declared('document', d.id, 'theme.colors.accent1'));
  assert.deepEqual(source(r, 'document', d.id, 'theme.fonts.heading'), builtin('defaults.theme.fonts.heading'));
  for (const [kind, id] of [['master', 'red'], ['layout', 'layout-0'], ['layout', 'layout-2'], ['slide', 'page-0'], ['slide', 'page-2']]) {
    assert.deepEqual(source(r, kind, id, 'theme.colors.accent1'), declared('master', 'red', 'masters.0.theme.colors.accent1'));
    assert.deepEqual(source(r, kind, id, 'theme.fonts.body'), declared('master', 'red', 'masters.0.theme.fonts.body'));
    assert.deepEqual(source(r, kind, id, 'theme.fonts.heading'), builtin('defaults.theme.fonts.heading'));
  }
  assert.deepEqual(source(r, 'slide', 'page-1', 'theme.name'), declared('master', 'blue', 'masters.1.theme.name'));
  assert.deepEqual(source(r, 'slide', 'page-3', 'theme.fonts.body'), declared('document', d.id, 'theme.fonts.body'));
  assert.deepEqual(source(r, 'layout', 'layout-blank', 'theme.colors.accent1'), declared('document', d.id, 'theme.colors.accent1'));
});

test('equal values retain distinct declaration origins, rather than inferred ownership by value', () => {
  const d = createPresentation({ id: 'equal', theme: { colors: { accent1: '123456' } }, masters: [{ id: 'a', name: 'A', theme: { colors: { accent1: '123456' } } }, { id: 'b', name: 'B' }] });
  for (const id of ['a', 'b']) { d.defineSlideLayout({ id, name: id, masterId: id }); d.addSlide({ id: `page-${id}`, layoutId: id }); }
  const r = resolvePresentationInheritance(d);
  assert.equal(r.presentation.masters[0].theme.colors.accent1, r.presentation.masters[1].theme.colors.accent1);
  assert.deepEqual(source(r, 'slide', 'page-a', 'theme.colors.accent1'), declared('master', 'a', 'masters.0.theme.colors.accent1'));
  assert.deepEqual(source(r, 'slide', 'page-b', 'theme.colors.accent1'), declared('document', 'equal', 'theme.colors.accent1'));
});

test('background paint selection does not merge properties from a lower-priority paint', () => {
  const d = createPresentation({ masters: [{ id: 'm', name: 'M', background: { type: 'solid', color: 'FF0000', opacity: 0.25 } }] });
  d.defineSlideLayout({ id: 'l', name: 'L', masterId: 'm', background: { type: 'solid', color: { theme: 'accent1' } } });
  d.addSlide({ id: 'inherit', layoutId: 'l' });
  d.addSlide({ id: 'local', layoutId: 'l', background: { type: 'solid', color: '000000', opacity: 0 } });
  d.addSlide({ id: 'clear', layoutId: 'l', background: { type: 'none' } });
  const r = resolvePresentationInheritance(d);
  assert.equal(r.presentation.slides[0].background.opacity, 1);
  assert.deepEqual(source(r, 'slide', 'inherit', 'background'), declared('layout', 'l', 'layouts.0.background'));
  assert.deepEqual(source(r, 'slide', 'inherit', 'background.color'), declared('layout', 'l', 'layouts.0.background.color'));
  assert.deepEqual(source(r, 'slide', 'inherit', 'background.opacity'), builtin('defaults.paint.opacity'));
  assert.deepEqual(source(r, 'slide', 'local', 'background.opacity'), declared('slide', 'local', 'slides.1.background.opacity'));
  assert.equal(r.presentation.slides[1].background.opacity, 0);
  assert.deepEqual(source(r, 'slide', 'clear', 'background.type'), declared('slide', 'clear', 'slides.2.background.type'));
  assert.deepEqual(r.origins.filter(row => row.target.kind === 'slide' && row.target.id === 'clear' && row.property.startsWith('background')).map(row => row.property), ['background', 'background.type']);
});

test('missing master paint reports the built-in fallback and never a fictitious theme background declaration', () => {
  const d = createPresentation({ masters: [{ id: 'm', name: 'M' }] });
  d.defineSlideLayout({ id: 'l', name: 'L', masterId: 'm' }); d.addSlide({ id: 's', layoutId: 'l' });
  const r = resolvePresentationInheritance(d);
  for (const [kind, id] of [['master', 'm'], ['layout', 'l'], ['slide', 's']]) {
    assert.deepEqual(source(r, kind, id, 'background'), builtin('defaults.background'));
    assert.deepEqual(source(r, kind, id, 'background.opacity'), builtin('defaults.background.opacity'));
  }
  assert.equal(r.presentation.slides[0].backgroundSource, 'theme'); // Legacy enum retained.
});

test('resetting shared definitions changes origins on the next resolve without mutating old reports', () => {
  const d = multiMasterDocument(createPresentation), before = resolvePresentationInheritance(d), old = JSON.stringify(before);
  d.replaceSlideMaster('red', { name: 'Resume document' });
  let after = resolvePresentationInheritance(d);
  for (const id of ['page-0', 'page-2']) {
    assert.deepEqual(source(after, 'slide', id, 'theme.colors.accent1'), declared('document', d.id, 'theme.colors.accent1'));
    assert.deepEqual(source(after, 'slide', id, 'background'), builtin('defaults.background'));
  }
  const priorLayout = d.layouts.find(l => l.id === 'layout-0');
  d.replaceSlideLayout('layout-0', { name: 'Blue', masterId: 'blue', placeholders: [...priorLayout.placeholders] });
  after = resolvePresentationInheritance(d);
  assert.deepEqual(source(after, 'slide', 'page-0', 'theme.colors.accent1'), declared('master', 'blue', 'masters.1.theme.colors.accent1'));
  assert.deepEqual(source(after, 'slide', 'page-2', 'theme.colors.accent1'), declared('document', d.id, 'theme.colors.accent1'));
  assert.equal(JSON.stringify(before), old);
});

test('report is an immutable snapshot while ordinary IR stays detached and mutable', () => {
  const d = multiMasterDocument(createPresentation), r = resolvePresentationInheritance(d), originBytes = JSON.stringify(r.origins);
  r.presentation.theme.colors.accent1 = 'FFFFFF'; r.presentation.masters[0].theme.colors.accent1 = '111111';
  r.presentation.layouts.find(l => l.id === 'layout-0').background.type = 'none';
  assert.equal(d.theme.colors.accent1, '112233'); assert.equal(d.masters[0].theme.colors.accent1, 'CC2200');
  assert.equal(JSON.stringify(r.origins), originBytes);
  assert.deepEqual(resolvePresentationInheritance(d).presentation, normalizePresentation(d));
});

test('authored default layout follows its real owner, including slides with omitted layout IDs', () => {
  const d = createPresentation({ masters: [{ id: 'm', name: 'M', background: { type: 'none' }, theme: { fonts: { body: 'Georgia' } } }] });
  d.defineSlideLayout({ id: 'layout-blank', name: 'Authored blank', masterId: 'm' }); d.addSlide({ id: 's' });
  const r = resolvePresentationInheritance(d);
  assert.equal(r.presentation.layouts.length, 1); assert.equal(r.presentation.slides[0].background.type, 'none');
  assert.deepEqual(source(r, 'slide', 's', 'theme.fonts.body'), declared('master', 'm', 'masters.0.theme.fonts.body'));
  assert.deepEqual(source(r, 'slide', 's', 'background'), declared('master', 'm', 'masters.0.background'));
  const keys = r.origins.map(row => JSON.stringify([row.target.kind, row.target.id, row.property])); assert.equal(new Set(keys).size, keys.length);
});

test('invalid input rejects with the same diagnostics and legacy custom documents need no extra interface', () => {
  const invalid = createPresentation(); invalid.defineSlideLayout({ id: 'l', name: 'L', masterId: 'missing' });
  let expected; try { normalizePresentation(invalid); } catch (error) { expected = error.diagnostics; }
  assert.throws(() => resolvePresentationInheritance(invalid), error => { assert.equal(error.name, 'PresentationValidationError'); assert.deepEqual(error.diagnostics, expected); return true; });
  const d = createPresentation({ id: 'legacy' }); d.addSlide({ id: 's' });
  const custom = Object.create(d); Object.defineProperty(custom, 'masters', { value: undefined });
  assert.deepEqual(resolvePresentationInheritance(custom), resolvePresentationInheritance(d));
  assert.equal(Object.hasOwn(resolvePresentationInheritance(custom).presentation, 'masters'), false);
});

test('every authored theme/background leaf resolves to the selected value and every theme role is covered once', () => {
  const d = multiMasterDocument(createPresentation);
  const local = d.addSlide({ id: 'local', layoutId: 'layout-0', background: { type: 'solid', color: '13579B', opacity: 0 } });
  const r = resolvePresentationInheritance(d), n = r.presentation;
  const read = (value, path) => path.split('.').reduce((current, key) => current[key], value);
  const themeFor = target => {
    if (target.kind === 'document') return n.theme;
    if (target.kind === 'master') return n.masters.find(m => m.id === target.id).theme;
    const layoutId = target.kind === 'layout' ? target.id : n.slides.find(s => s.id === target.id).layoutId;
    return n.layouts.find(l => l.id === layoutId).theme ?? n.theme;
  };
  const backgroundFor = target => {
    const collection = target.kind === 'master' ? n.masters : target.kind === 'layout' ? n.layouts : n.slides;
    return collection.find(value => value.id === target.id).background;
  };
  const roles = ['background1', 'background2', 'text1', 'text2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hyperlink', 'followedHyperlink'];
  const targets = new Map(r.origins.filter(row => row.property.startsWith('theme.')).map(row => [JSON.stringify(row.target), row.target]));
  for (const target of targets.values()) {
    const themeRows = r.origins.filter(row => row.target.kind === target.kind && row.target.id === target.id && row.property.startsWith('theme.'));
    assert.deepEqual(themeRows.map(row => row.property).sort(), ['theme.name', ...roles.map(role => `theme.colors.${role}`), 'theme.fonts.heading', 'theme.fonts.body'].sort());
  }
  for (const row of r.origins) {
    if (!row.property.startsWith('theme.') && !row.property.startsWith('background')) continue;
    if (row.source.kind === 'default' || row.property === 'background') continue;
    const expected = row.property.startsWith('theme.') ? read(themeFor(row.target), row.property.slice(6)) : read(backgroundFor(row.target), row.property.slice(11));
    assert.deepEqual(read(d, row.source.path), expected, JSON.stringify(row));
  }
  assert.equal(source(r, 'slide', local.id, 'background.opacity').kind, 'slide');
});

test('source paths follow current collection positions while logical IDs survive slide moves and punctuation', () => {
  const d = createPresentation({ id: 'doc.with/slash', masters: [{ id: 'master.with/slash', name: 'M', theme: { colors: { accent1: '123456' } } }] });
  d.defineSlideLayout({ id: 'layout.with/slash', name: 'L', masterId: 'master.with/slash' });
  d.addSlide({ id: 'a.with/slash', layoutId: 'layout.with/slash', background: { type: 'solid', color: '123456', opacity: 0 } }); d.addSlide({ id: 'b' });
  const before = resolvePresentationInheritance(d);
  d.moveSlide('a.with/slash', 1);
  const after = resolvePresentationInheritance(d);
  assert.deepEqual(source(before, 'slide', 'a.with/slash', 'background.opacity'), declared('slide', 'a.with/slash', 'slides.0.background.opacity'));
  assert.deepEqual(source(after, 'slide', 'a.with/slash', 'background.opacity'), declared('slide', 'a.with/slash', 'slides.1.background.opacity'));
  assert.deepEqual(source(after, 'slide', 'a.with/slash', 'theme.colors.accent1'), declared('master', 'master.with/slash', 'masters.0.theme.colors.accent1'));
});

test('custom theme records attribute only the enumerable declarations used by the value merge', () => {
  const d = createPresentation({ id: 'custom' }), custom = Object.create(d);
  const colors = Object.defineProperty({ accent2: '13579B' }, 'accent1', { value: '123456', enumerable: false });
  const fonts = Object.defineProperty({}, 'body', { value: 'Georgia', enumerable: false });
  let names = 0;
  const theme = { colors, fonts, get name() { names++; return 'Custom theme'; } };
  Object.defineProperty(custom, 'theme', { value: theme });
  const r = resolvePresentationInheritance(custom);
  assert.equal(names, 1);
  assert.deepEqual(source(r, 'document', d.id, 'theme.colors.accent1'), builtin('defaults.theme.colors.accent1'));
  assert.deepEqual(source(r, 'document', d.id, 'theme.fonts.body'), builtin('defaults.theme.fonts.body'));
  assert.deepEqual(source(r, 'document', d.id, 'theme.colors.accent2'), declared('document', d.id, 'theme.colors.accent2'));
  assert.equal(r.presentation.theme.colors.accent2, '13579B'); assert.notEqual(r.presentation.theme.fonts.body, 'Georgia');
});
