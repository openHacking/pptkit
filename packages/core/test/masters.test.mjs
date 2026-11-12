import assert from 'node:assert/strict';
import test from 'node:test';
import { createPresentation, normalizePresentation, validatePresentation } from '../dist/index.js';
import { multiMasterDocument } from '../../../tests/fixtures/multi-master.mjs';

test('master definitions are captured and detached, while layouts bind distinct same-key theme/background domains', () => {
  const input = [{ id: 'one', name: 'One', theme: { colors: { accent1: '123456' } } }];
  const owned = createPresentation({ masters: input }); input[0].theme.colors.accent1 = '999999'; input.push({ id: 'other', name: 'Other' });
  assert.equal(owned.masters.length, 1); assert.equal(owned.masters[0].theme.colors.accent1, '123456'); assert.ok(Object.isFrozen(owned.masters[0].theme.colors)); assert.throws(() => { owned.masters = []; }, TypeError);
  const document = multiMasterDocument(createPresentation), n = normalizePresentation(document);
  assert.equal(n.masters.length, 3); assert.deepEqual(n.layouts.filter(l => l.masterId).map(l => l.theme.colors.accent1), ['CC2200', '0044CC', 'CC2200']);
  assert.deepEqual(n.slides.map(s => s.backgroundSource), ['master', 'master', 'master', 'theme']);
  n.layouts[1].theme.colors.accent1 = 'FFFFFF'; assert.equal(document.masters[0].theme.colors.accent1, 'CC2200'); assert.equal(n.masters[0].theme.colors.accent1, 'CC2200');
  assert.equal(document.duplicateSlide('page-0').layoutId, 'layout-0');
});

test('legacy defaults omit all new IR fields and custom document implementations need no new methods', () => {
  const a = createPresentation({ id: 'same' }); a.addSlide({ id: 'slide' });
  const custom = Object.create(a); Object.defineProperty(custom, 'masters', { value: undefined });
  assert.deepEqual(normalizePresentation(a), normalizePresentation(custom)); assert.equal(Object.hasOwn(normalizePresentation(a), 'masters'), false);
});

test('missing references and invalid/duplicate master declarations fail normalization with paths, without cross-domain fallback', () => {
  const d = createPresentation({ masters: [{ id: 'a', name: 'A' }, { id: 'a', name: '', theme: { colors: { accent1: 'bad' }, fonts: { body: '' } } }] });
  d.defineSlideLayout({ id: 'missing', name: 'Missing', masterId: 'nope' });
  const codes = validatePresentation(d).map(d => d.code);
  for (const code of ['duplicate-master-id', 'invalid-master-name', 'invalid-theme-color', 'invalid-font-family', 'missing-layout-master']) assert.ok(codes.includes(code));
  assert.throws(() => normalizePresentation(d), e => e.name === 'PresentationValidationError' && e.diagnostics.some(d => d.path === 'layouts.0.masterId'));
});

test('local none and zero-opacity backgrounds override master paint without truthiness fallback', () => {
  const d = createPresentation({ masters: [{ id: 'm', name: 'M', background: { type: 'solid', color: 'FF0000' } }] });
  d.defineSlideLayout({ id: 'a', name: 'A', masterId: 'm', background: { type: 'none' } });
  d.defineSlideLayout({ id: 'b', name: 'B', masterId: 'm' });
  d.addSlide({ layoutId: 'a' }); d.addSlide({ layoutId: 'b', background: { type: 'solid', color: '000000', opacity: 0 } });
  const n = normalizePresentation(d); assert.equal(n.slides[0].background.type, 'none'); assert.equal(n.slides[0].backgroundSource, 'layout');
  assert.equal(n.slides[1].background.opacity, 0); assert.equal(n.slides[1].backgroundSource, 'slide');
});

test('unsupported master graph declarations and malformed artwork and malformed theme values fail explicitly', () => {
  for (const masters of [null, {}, [null], [{ id: 'x', name: 'X', parentId: 'x' }], [{ id: 'x', name: 'X', elements: {} }]]) assert.throws(() => createPresentation({ masters }), TypeError);
  const hidden = Object.defineProperty({ id: 'x', name: 'X' }, 'parentId', { value: 'x' });
  assert.throws(() => createPresentation({ masters: [hidden] }), TypeError);
  for (const theme of [[], 'wrong', { name: 2 }, { fonts: { body: undefined } }, { unknown: true }, { colors: null }, { fonts: [] }, { colors: { unsupported: '123456' } }]) {
    const d = createPresentation({ masters: [{ id: 'x', name: 'X', theme }] });
    assert.throws(() => normalizePresentation(d), e => e.name === 'PresentationValidationError');
  }
});
