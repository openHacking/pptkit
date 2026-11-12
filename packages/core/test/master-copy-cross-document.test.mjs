import test from 'node:test';
import assert from 'node:assert/strict';
import { createPresentation, normalizePresentation } from '../dist/index.js';

const box = { x: 0, y: 0, width: 100, height: 40 };
const policies = { id: 'copy', theme: 'source-declarations', assets: { deduplicate: 'content', external: 'reject' } };

test('equivalent master reuse preserves references, resources and allocator state', () => {
  const source = createPresentation(), target = createPresentation();
  source.registerAsset({ id: 'img', kind: 'image', source: { type: 'bytes', value: Uint8Array.of(1, 2) } });
  source.defineSlideMaster({ id: 'master', name: 'Shared', theme: { colors: { accent1: '123456' } }, elements: [
    { id: 'picture', type: 'image', assetId: 'img', box },
    { id: 'edge', type: 'connector', start: { elementId: 'picture' }, end: { x: 120, y: 20 } },
  ] });
  const first = target.copySlideMasterFrom(source, 'master', policies);
  const reused = target.copySlideMasterFrom(source, 'master', { ...policies, id: 'next', reuseEquivalent: true });
  assert.equal(reused.master, first.master);
  assert.deepEqual(reused.elementIds, first.elementIds);
  assert.equal(reused.assets[0].reused, true);
  assert.equal(target.masters.length, 1); assert.equal(target.assets.length, 1);
  const fresh = target.copySlideMasterFrom(source, 'master', { ...policies, id: 'next' });
  assert.equal(fresh.elementIds.picture, 'element-3');
  assert.equal(target.copySlideMasterFrom(source, 'master', { ...policies, id: 'third', reuseEquivalent: true }).master, first.master);
});

test('master reuse distinguishes declarations, artwork and remapped actions', () => {
  const source = createPresentation({ theme: { colors: { accent1: '123456' } } }), target = createPresentation({ theme: { colors: { accent1: '123456' } } });
  source.addSlide({ id: 'source-page' }); target.addSlide({ id: 'a' }); target.addSlide({ id: 'b' });
  source.defineSlideMaster({ id: 'master', name: 'Shared', theme: { colors: { accent1: '123456' } }, elements: [{ id: 'shape', type: 'shape', shape: 'rect', box, action: { type: 'slide', slideId: 'source-page' } }] });
  const copy = (id, theme, page) => target.copySlideMasterFrom(source, 'master', { ...policies, id, theme, reuseEquivalent: true, slideLinks: { 'source-page': page } });
  const first = copy('first', 'source-declarations', 'a');
  assert.equal(copy('different-action', 'source-declarations', 'b').master.id, 'different-action');
  assert.equal(copy('different-inheritance', 'destination', 'a').master.id, 'different-inheritance');
  assert.equal(copy('reuse', 'source-declarations', 'a').master, first.master);
  assert.throws(() => target.copySlideMasterFrom(source, 'master', { ...policies, id: 'bad', reuseEquivalent: 'yes' }), /boolean/);
});

test('master reuse rejects a reentrant change to the previously matching master', () => {
  const source = createPresentation({ theme: { colors: { accent1: '123456' } } }), target = createPresentation();
  source.defineSlideMaster({ id: 'master', name: 'Shared', elements: [{ id: 'shape', type: 'shape', shape: 'rect', box }] });
  const opts = { ...policies, theme: 'source-effective' };
  const first = target.copySlideMasterFrom(source, 'master', opts);
  const hostile = Object.create(source, { theme: { get() {
    target.replaceSlideMaster(first.master.id, { ...first.master, name: 'Changed independently' });
    return source.theme;
  } } });
  assert.throws(() => target.copySlideMasterFrom(hostile, 'master', { ...opts, id: 'next', reuseEquivalent: true }), /changed/);
  assert.equal(target.masters.length, 1);
  assert.equal(target.masters[0].name, 'Changed independently');
  assert.equal(target.copySlideMasterFrom(source, 'master', { ...opts, id: 'next', reuseEquivalent: true }).elementIds.shape, 'element-2');
});

test('master reuse preserves the distinction between absent and explicit empty artwork', () => {
  const source = createPresentation(), target = createPresentation();
  source.defineSlideMaster({ id: 'empty', name: 'Master', elements: [] });
  source.defineSlideMaster({ id: 'absent', name: 'Master' });
  const empty = target.copySlideMasterFrom(source, 'empty', { ...policies, id: 'empty-copy', reuseEquivalent: true });
  const absent = target.copySlideMasterFrom(source, 'absent', { ...policies, id: 'absent-copy', reuseEquivalent: true });
  assert.equal(empty.master.id, 'empty-copy');
  assert.equal(absent.master.id, 'absent-copy');
  assert.ok(Object.hasOwn(empty.master, 'elements'));
  assert.equal(Object.hasOwn(absent.master, 'elements'), false);
  assert.equal(target.copySlideMasterFrom(source, 'absent', { ...policies, id: 'unused', reuseEquivalent: true }).master, absent.master);
});

test('master reuse creates a separate master when matching bytes have different accessibility metadata', () => {
  const source = createPresentation(), target = createPresentation();
  source.registerAsset({ id: 'image', kind: 'image', source: { type: 'bytes', value: Uint8Array.of(7) }, accessibility: { description: 'First' } });
  source.defineSlideMaster({ id: 'master', name: 'Master', elements: [{ id: 'picture', type: 'image', assetId: 'image', box }] });
  const first = target.copySlideMasterFrom(source, 'master', policies);
  source.replaceAsset('image', { kind: 'image', source: { type: 'bytes', value: Uint8Array.of(7) }, accessibility: { description: 'Second' } });
  const changed = target.copySlideMasterFrom(source, 'master', { ...policies, id: 'second', reuseEquivalent: true });
  assert.equal(changed.master.id, 'second');
  assert.equal(changed.assets[0].reused, false);
  assert.notEqual(changed.assets[0].destinationId, first.assets[0].destinationId);
  assert.equal(target.assets.length, 2);
  assert.equal(target.getAsset(first.assets[0].destinationId).accessibility.description, 'First');
  assert.equal(target.getAsset(changed.assets[0].destinationId).accessibility.description, 'Second');
});

test('master copy captures mapping values before reentrant source access', () => {
  const source = createPresentation({ textStylePresets: { old: {} } }), dest = createPresentation({ textStylePresets: { first: {}, second: {} } });
  source.defineSlideMaster({ id: 'master', name: 'Master', elements: [{ id: 'text', type: 'text', box, textStylePreset: 'old', content: 'A' }] });
  let calls = 0;
  const getter = Object.defineProperty({}, 'old', { enumerable: true, get() { calls++; return 'first'; } });
  for (const field of ['slideLinks', 'textPresets']) {
    assert.throws(() => dest.copySlideMasterFrom(source, 'master', { ...policies, [field]: getter }), TypeError);
    assert.equal(dest.masters.length, 0); assert.equal(dest.assets.length, 0);
  }
  assert.equal(calls, 0);
  const mapping = { old: 'first' };
  const changing = Object.create(source, { assets: { get() { mapping.old = 'second'; return source.assets; } } });
  const result = dest.copySlideMasterFrom(changing, 'master', { ...policies, textPresets: mapping });
  assert.equal(mapping.old, 'second'); assert.equal(result.master.elements[0].textStylePreset, 'first');
});

test('copies master artwork, assets and explicit local theme declarations', () => {
  const source = createPresentation(), dest = createPresentation();
  source.registerAsset({ id: 'image', kind: 'image', source: { type: 'bytes', value: Uint8Array.of(1, 2) } });
  source.defineSlideMaster({ id: 'master', name: 'Master', theme: { colors: { accent1: 'FF0000' } }, elements: [
    { id: 'edge', type: 'connector', start: { elementId: 'image' }, end: { x: 100, y: 40 } },
    { id: 'image', type: 'image', assetId: 'image', box },
  ] });
  const copy = dest.copySlideMasterFrom(source, 'master', policies);
  assert.equal(copy.master.theme.colors.accent1, 'FF0000');
  assert.equal(copy.master.elements[0].start.elementId, copy.elementIds.image);
  assert.equal(copy.master.elements[1].assetId, copy.assets[0].destinationId);
  source.replaceAsset('image', { kind: 'image', source: { type: 'bytes', value: Uint8Array.of(9) } });
  assert.equal(dest.assets[0].source.value, 'AQI=');
  assert.ok(Object.isFrozen(copy.master));
  const adopted = dest.copySlideMasterFrom(source, 'master', { ...policies, id: 'adopted', theme: 'destination' });
  assert.equal(adopted.master.theme, undefined);
});

test('master link mappings and invalid candidates fail without partial commit', () => {
  const source = createPresentation(), dest = createPresentation();
  source.registerAsset({ id: 'image', kind: 'image', source: { type: 'bytes', value: Uint8Array.of(1) } });
  source.addSlide({ id: 'source-page' });
  source.defineSlideMaster({ id: 'master', name: 'Master', elements: [{ id: 'image', type: 'image', assetId: 'image', box, action: { type: 'slide', slideId: 'source-page' } }] });
  assert.throws(() => dest.copySlideMasterFrom(source, 'master', policies), /mapping/);
  assert.equal(dest.assets.length, 0); assert.equal(dest.masters.length, 0);
  dest.addSlide({ id: 'target' });
  const copy = dest.copySlideMasterFrom(source, 'master', { ...policies, slideLinks: { 'source-page': 'target' } });
  assert.equal(copy.master.elements[0].action.slideId, 'target'); assert.equal(copy.elementIds.image, 'element-1');
  assert.throws(() => dest.copySlideMasterFrom(source, 'master', { ...policies, slideLinks: { 'source-page': 'target' } }), /Duplicate/);
  assert.equal(dest.masters.length, 1); assert.equal(dest.assets.length, 1);
});

test('effective source theme is materialized independently of the destination theme', () => {
  const source = createPresentation({ theme: { name: 'Source', colors: { accent1: '112233', accent2: '445566' }, fonts: { body: 'Source Body', heading: 'Source Heading' } } });
  source.defineSlideMaster({ id: 'master', name: 'Master', theme: { colors: { accent1: 'AABBCC' }, fonts: { heading: 'Master Heading' } } });
  const dest = createPresentation({ theme: { name: 'Destination', colors: { accent1: 'FFFFFF', accent2: '000000' }, fonts: { body: 'Destination Body' } } });
  const result = dest.copySlideMasterFrom(source, 'master', { ...policies, theme: 'source-effective' });
  const sourceTheme = normalizePresentation(source).masters[0].theme;
  const copiedTheme = normalizePresentation(dest).masters[0].theme;
  assert.deepEqual(copiedTheme, sourceTheme);
  assert.equal(result.master.theme.colors.accent1, 'AABBCC');
  assert.equal(result.master.theme.colors.accent2, '445566');
  assert.equal(result.master.theme.fonts.body, 'Source Body');
  assert.equal(result.master.theme.fonts.heading, 'Master Heading');
  assert.ok(Object.isFrozen(result.master.theme.colors));
});

test('cross-document master theme policies retain distinct source and destination inheritance', () => {
  const source = createPresentation({ theme: { colors: { accent1: '112233', accent2: '445566' }, fonts: { body: 'Source Body', heading: 'Source Heading' } } });
  source.defineSlideMaster({ id: 'master', name: 'Master', theme: { colors: { accent1: 'ABCDEF' }, fonts: { heading: 'Master Heading' } } });
  const destination = createPresentation({ theme: { colors: { accent1: '010203', accent2: 'A0B0C0' }, fonts: { body: 'Destination Body', heading: 'Destination Heading' } } });
  for (const theme of ['destination', 'source-declarations', 'source-effective']) {
    destination.copySlideMasterFrom(source, 'master', { ...policies, id: theme, theme });
  }
  const themes = Object.fromEntries(normalizePresentation(destination).masters.map(master => [master.id, master.theme]));
  assert.deepEqual([themes.destination.colors.accent1, themes.destination.colors.accent2, themes.destination.fonts.body, themes.destination.fonts.heading], ['010203', 'A0B0C0', 'Destination Body', 'Destination Heading']);
  assert.deepEqual([themes['source-declarations'].colors.accent1, themes['source-declarations'].colors.accent2, themes['source-declarations'].fonts.body, themes['source-declarations'].fonts.heading], ['ABCDEF', 'A0B0C0', 'Destination Body', 'Master Heading']);
  assert.deepEqual([themes['source-effective'].colors.accent1, themes['source-effective'].colors.accent2, themes['source-effective'].fonts.body, themes['source-effective'].fonts.heading], ['ABCDEF', '445566', 'Source Body', 'Master Heading']);
  source.replaceSlideMaster('master', { id: 'master', name: 'Changed', theme: { colors: { accent1: 'FFFFFF' } } });
  assert.deepEqual(Object.fromEntries(normalizePresentation(destination).masters.map(master => [master.id, master.theme])), themes);
});

test('master copy detects reentrant destination changes and preserves separate edits', () => {
  const source = createPresentation(), dest = createPresentation();
  source.registerAsset({ id: 'image', kind: 'image', source: { type: 'bytes', value: Uint8Array.of(1) } });
  source.defineSlideMaster({ id: 'master', name: 'Master', elements: [{ id: 'image', type: 'image', assetId: 'image', box }] });
  const hostile = Object.create(source, { assets: { get() { dest.registerAsset({ id: 'independent', kind: 'image', source: { type: 'bytes', value: Uint8Array.of(9) } }); return source.assets; } } });
  assert.throws(() => dest.copySlideMasterFrom(hostile, 'master', policies), /changed/);
  assert.deepEqual(dest.assets.map(item => item.id), ['independent']); assert.equal(dest.masters.length, 0);
  const result = dest.copySlideMasterFrom(source, 'master', policies);
  assert.equal(result.elementIds.image, 'element-1');
});

test('master copy detects a reentrantly added empty target layout', () => {
  const source = createPresentation(), dest = createPresentation();
  source.defineSlideMaster({ id: 'master', name: 'Master' });
  const hostile = Object.create(source, { assets: { get() { dest.defineSlideLayout({ id: 'independent', name: 'Independent' }); return source.assets; } } });
  assert.throws(() => dest.copySlideMasterFrom(hostile, 'master', policies), /changed/);
  assert.equal(dest.masters.length, 0); assert.deepEqual(dest.layouts.map(layout => layout.id), ['independent']);
  assert.equal(dest.copySlideMasterFrom(source, 'master', policies).master.id, 'copy');
});
