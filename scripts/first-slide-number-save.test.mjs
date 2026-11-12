import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { chromium } from '@playwright/test';
import { createPresentation } from '../packages/core/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';
import { importPptx, inspectPptx } from '../packages/pptx-parser/dist/index.js';
import { importPptxFile, writePptxFile } from '../packages/pptx-parser/dist/node.js';
import { readGeneratedZip } from '../tests/support/read-generated-zip.mjs';
import { structureFixture, encodeStructure } from '../packages/pptx-parser/test/structure-fixture.mjs';

test('existing firstSlideNum writes one source attribute through three validated edit rounds', async () => {
  const deck = createPresentation({ firstSlideNumber: 1 }); deck.addSlide({ elements: [{ type: 'text', box: { x: 10, y: 10, width: 100, height: 30 }, content: [{ runs: [{ text: 'unchanged' }, { text: 'cache', field: { id: '38d9dc4a-2e99-e34a-a7a3-8e01bbc23a4a', kind: 'slide-number' } }] }] }] });
  let bytes = (await generatePptx(deck)).bytes, previous = '1';
  for (const value of [0, 17, 9999]) {
    const session = await importPptx(bytes, { projection: 'content', fieldTextProfile: 'slide-number-v1', firstSlideNumberSaveProfile: 'existing-v1' });
    try {
      assert.equal(session.document.firstSlideNumber, Number(previous)); session.document.setFirstSlideNumber(value);
      const plan = await session.prepareSave({ mode: 'edit' }); assert.equal(plan.status, 'ready', JSON.stringify(plan.report));
      assert.deepEqual(plan.plan.report.changes.map(change => change.firstSlideNumber), [{ before: Number(previous), after: value }]);
      assert.equal(plan.plan.report.changes[0].source.part, '/ppt/presentation.xml');
      const next = (await session.save(plan.plan)).bytes, before = readGeneratedZip(bytes), after = readGeneratedZip(next);
      assert.deepEqual([...before.keys()].sort(), [...after.keys()].sort());
      for (const [name, payload] of before) {
        if (name === 'ppt/presentation.xml') assert.equal(after.get(name).toString(), payload.toString().replace(`firstSlideNum="${previous}"`, `firstSlideNum="${value}"`));
        else assert.deepEqual(after.get(name), payload);
      }
      const inspection = await inspectPptx(next); assert.equal(inspection.firstSlideNumber, value); assert.equal(inspection.firstSlideNumberDeclared, true);
      bytes = next; previous = String(value);
    } finally { session.dispose(); }
  }
});

test('file adapter captures the number-save profile and writes a verified new file', async t => {
  const directory = await mkdtemp(path.join(tmpdir(), 'pptkit-number-save-')); t.after(() => rm(directory, { recursive: true, force: true }));
  const source = path.join(directory, 'source.pptx'), target = path.join(directory, 'target.pptx');
  const deck = createPresentation({ firstSlideNumber: 1 }); deck.addSlide(); await writeFile(source, (await generatePptx(deck)).bytes);
  const imported = await importPptxFile(source, { projection: 'content', firstSlideNumberSaveProfile: 'existing-v1' });
  try { imported.session.document.setFirstSlideNumber(42); const plan = await imported.session.prepareSave({ mode: 'edit' }); assert.equal(plan.status, 'ready'); await writePptxFile(imported.session, plan.plan, target, { mode: 'create' }); assert.equal((await inspectPptx(await readFile(target))).firstSlideNumber, 42); assert.equal((await inspectPptx(await readFile(source))).firstSlideNumber, 1); } finally { imported.session.dispose(); }
});

test('browser number-only preservation matches Node and blocks unsupported mixed edits', async t => {
  const deck = createPresentation({ firstSlideNumber: 17 }); deck.addSlide({ elements: [{ type: 'text', box: { x: 10, y: 10, width: 100, height: 30 }, content: 'before' }] });
  const bytes = (await generatePptx(deck)).bytes, session = await importPptx(bytes, { projection: 'content', firstSlideNumberSaveProfile: 'existing-v1' });
  let expected;
  try { session.document.setFirstSlideNumber(30); const plan = await session.prepareSave({ mode: 'edit' }); assert.equal(plan.status, 'ready'); expected = (await session.save(plan.plan)).bytes; } finally { session.dispose(); }
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
  for (const name of ['core', 'pptx-parser']) await page.addScriptTag({ content: await readFile(`packages/${name}/dist/browser/global.js`, 'utf8') });
  const actual = await page.evaluate(async input => { const s = await PPTKit.importPptx(new Uint8Array(input), { projection: 'content', firstSlideNumberSaveProfile: 'existing-v1' }); try { s.document.setFirstSlideNumber(30); const p = await s.prepareSave({ mode: 'edit' }), out = await s.save(p.plan); const slide = s.document.slides[0], element = slide.elements[0], changed = structuredClone(element); changed.content[0].runs[0].text = 'unsupported'; slide.replaceElement(element.id, changed); const mixed = await s.prepareSave({ mode: 'edit' }); return { bytes: Array.from(out.bytes), mixed: mixed.status }; } finally { s.dispose(); } }, Array.from(bytes));
  assert.deepEqual(new Uint8Array(actual.bytes), expected); assert.equal(actual.mixed, 'blocked'); assert.deepEqual(requests, []);
});

test('existing numeric save targets the selected MCE root in both dialects without changing inactive branches', async () => {
  for (const strict of [false, true]) {
    const fixture = structureFixture(strict), root = fixture.parts['deck/main.data'];
    const selected = root.replace('<p:presentation ', '<p:presentation firstSlideNum="+001" ');
    const inactive = root.replace('<p:presentation ', '<p:presentation firstSlideNum="7" ');
    fixture.parts['deck/main.data'] = `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:unsupported"><mc:Choice Requires="x">${inactive}</mc:Choice><mc:Fallback>${selected}</mc:Fallback></mc:AlternateContent>`;
    const bytes = encodeStructure(fixture), session = await importPptx(bytes, { projection: 'content', firstSlideNumberSaveProfile: 'existing-v1' });
    try {
      assert.equal(session.document.firstSlideNumber, 1); session.document.setFirstSlideNumber(17);
      const plan = await session.prepareSave({ mode: 'edit' }); assert.equal(plan.status, 'ready', JSON.stringify(plan.report));
      const output = (await session.save(plan.plan)).bytes, before = readGeneratedZip(bytes), after = readGeneratedZip(output);
      for (const [name, payload] of before) {
        if (name === 'deck/main.data') assert.equal(after.get(name).toString(), payload.toString().replace('firstSlideNum="+001"', 'firstSlideNum="17"'));
        else assert.deepEqual(after.get(name), payload);
      }
      const inspection = await inspectPptx(output); assert.equal(inspection.firstSlideNumber, 17); assert.equal(inspection.dialect, strict ? 'strict' : 'transitional');
    } finally { session.dispose(); }
  }
});

test('a nonnegative source value outside the PowerPoint range can be replaced with an admitted value', async () => {
  const fixture = structureFixture(); fixture.parts['deck/main.data'] = fixture.parts['deck/main.data'].replace('<p:presentation ', '<p:presentation firstSlideNum="10000" ');
  const bytes = encodeStructure(fixture), session = await importPptx(bytes, { projection: 'content', firstSlideNumberSaveProfile: 'existing-v1' });
  try {
    assert.equal(session.document.firstSlideNumber, 10000); session.document.setFirstSlideNumber(17);
    const plan = await session.prepareSave({ mode: 'edit' }); assert.equal(plan.status, 'ready');
    assert.deepEqual(plan.plan.report.changes[0].firstSlideNumber, { before: 10000, after: 17 });
    assert.equal((await inspectPptx((await session.save(plan.plan)).bytes)).firstSlideNumber, 17);
  } finally { session.dispose(); }
});

test('source-number save preserves UTF-16 bytes outside the value span', async () => {
  for (const endian of ['le', 'be']) {
    const fixture = structureFixture(), sourceText = fixture.parts['deck/main.data'].replace('<p:presentation ', "<p:presentation firstSlideNum = '+001' ");
    const encode = text => { const bytes = Buffer.from(text, 'utf16le'); if (endian === 'be') bytes.swap16(); return Buffer.concat([Buffer.from(endian === 'le' ? [0xff, 0xfe] : [0xfe, 0xff]), bytes]); };
    fixture.parts['deck/main.data'] = encode(sourceText);
    const bytes = encodeStructure(fixture), session = await importPptx(bytes, { projection: 'content', firstSlideNumberSaveProfile: 'existing-v1' });
    try {
      session.document.setFirstSlideNumber(42); const plan = await session.prepareSave({ mode: 'edit' }); assert.equal(plan.status, 'ready');
      const output = (await session.save(plan.plan)).bytes, parts = readGeneratedZip(output);
      assert.deepEqual(parts.get('deck/main.data'), encode(sourceText.replace("firstSlideNum = '+001'", "firstSlideNum = '42'")));
      assert.equal((await inspectPptx(output)).firstSlideNumber, 42);
    } finally { session.dispose(); }
  }
});

test('number profile blocks insertion, clearing, out-of-range, mixed edits and stale plans', async () => {
  const source = async value => { const deck = createPresentation(value === undefined ? {} : { firstSlideNumber: value }); deck.addSlide({ elements: [{ type: 'text', box: { x: 10, y: 10, width: 100, height: 30 }, content: 'before' }] }); return (await generatePptx(deck)).bytes; };
  const absent = await importPptx(await source(undefined), { projection: 'content', firstSlideNumberSaveProfile: 'existing-v1' });
  try { absent.document.setFirstSlideNumber(3); assert.equal((await absent.prepareSave({ mode: 'edit' })).status, 'blocked'); } finally { absent.dispose(); }
  const session = await importPptx(await source(1), { projection: 'content', firstSlideNumberSaveProfile: 'existing-v1' });
  try {
    const cancelled = new AbortController(); cancelled.abort();
    session.document.setFirstSlideNumber(2); await assert.rejects(session.prepareSave({ mode: 'edit', signal: cancelled.signal }), e => e.code === 'operation-aborted');
    session.document.setFirstSlideNumber(); assert.equal((await session.prepareSave({ mode: 'edit' })).status, 'blocked');
    session.document.setFirstSlideNumber(10000); assert.equal((await session.prepareSave({ mode: 'edit' })).status, 'blocked');
    session.document.setFirstSlideNumber(2); const slide = session.document.slides[0], element = slide.elements[0], changed = structuredClone(element); changed.content[0].runs[0].text = 'mixed'; slide.replaceElement(element.id, changed);
    assert.equal((await session.prepareSave({ mode: 'edit' })).status, 'blocked'); slide.replaceElement(element.id, element);
    const plan = await session.prepareSave({ mode: 'edit' }); assert.equal(plan.status, 'ready'); session.document.setFirstSlideNumber(3);
    await assert.rejects(session.save(plan.plan), e => e.code === 'save-snapshot-stale');
  } finally { session.dispose(); }
  await assert.rejects(importPptx(await source(1), { projection: 'structure', firstSlideNumberSaveProfile: 'existing-v1' }), TypeError);
  await assert.rejects(importPptx(await source(1), { projection: 'content', firstSlideNumberSaveProfile: 'existing-v1', editSaveProfile: 'text-runs-v3' }), TypeError);
  await assert.rejects(importPptx(await source(1), { projection: 'content', firstSlideNumberSaveProfile: 'existing-v1', pictureSourceProfile: 'embedded-v1' }), TypeError);
  await assert.rejects(importPptx(await source(1), { projection: 'content', firstSlideNumberSaveProfile: 'unknown' }), TypeError);
});


test('attribute-v2 inserts, replaces and clears explicit start numbers across reopen cycles', async () => {
  const deck = createPresentation(); deck.addSlide(); let bytes = (await generatePptx(deck)).bytes;
  let previous;
  for (const value of [1, 0, undefined, 17, undefined]) {
    const session = await importPptx(bytes, { projection: 'content', firstSlideNumberSaveProfile: 'attribute-v2' });
    try {
      assert.equal(session.document.firstSlideNumber, previous); session.document.setFirstSlideNumber(value);
      const prepared = await session.prepareSave({ mode: 'edit' }); assert.equal(prepared.status, 'ready', JSON.stringify(prepared.report));
      assert.deepEqual(prepared.plan.report.changes[0].firstSlideNumber, { ...(previous !== undefined ? { before: previous } : {}), ...(value !== undefined ? { after: value } : {}) });
      const next = (await session.save(prepared.plan)).bytes, before = readGeneratedZip(bytes), after = readGeneratedZip(next);
      for (const [name, payload] of before) {
        if (name !== 'ppt/presentation.xml') assert.deepEqual(after.get(name), payload);
        else {
          const original = payload.toString(), result = after.get(name).toString();
          if (previous === undefined) assert.equal(result.replace(` firstSlideNum="${value}"`, ''), original);
          else if (value === undefined) assert.equal(result, original.replace(`firstSlideNum="${previous}"`, ''));
          else assert.equal(result, original.replace(`firstSlideNum="${previous}"`, `firstSlideNum="${value}"`));
        }
      }
      const inspection = await inspectPptx(next); assert.equal(Boolean(inspection.firstSlideNumberDeclared), value !== undefined); assert.equal(inspection.firstSlideNumber, value ?? 1);
      previous = value; bytes = next;
    } finally { session.dispose(); }
  }
});

test('attribute-v2 edits only selected MCE branch and preserves UTF-16 source', async () => {
  for (const strict of [false, true]) for (const endian of ['le', 'be']) {
    const fixture = structureFixture(strict), root = fixture.parts['deck/main.data'];
    const source = `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:unsupported"><mc:Choice Requires="x">${root.replace('<p:presentation ', '<p:presentation firstSlideNum="7" ')}</mc:Choice><mc:Fallback>${root}</mc:Fallback></mc:AlternateContent>`;
    const encode = text => { const data = Buffer.from(text, 'utf16le'); if (endian === 'be') data.swap16(); return Buffer.concat([Buffer.from(endian === 'le' ? [255,254] : [254,255]), data]); };
    fixture.parts['deck/main.data'] = encode(source); let bytes = encodeStructure(fixture);
    for (const value of [1, undefined]) {
      const session = await importPptx(bytes, { projection: 'content', firstSlideNumberSaveProfile: 'attribute-v2' });
      try {
        session.document.setFirstSlideNumber(value); const plan = await session.prepareSave({ mode: 'edit' }); assert.equal(plan.status, 'ready', JSON.stringify(plan.report));
        const next = (await session.save(plan.plan)).bytes, before = readGeneratedZip(bytes), after = readGeneratedZip(next);
        const old = before.get('deck/main.data'), current = after.get('deck/main.data');
        // Decode explicitly for both byte orders before checking the exact lexical change.
        const decode = data => { const copy = Buffer.from(data.subarray(2)); if (endian === 'be') copy.swap16(); return copy.toString('utf16le'); };
        const text = decode(old), changed = decode(current);
        if (value === 1) assert.equal(changed.replace(' firstSlideNum="1"', ''), text);
        else assert.equal(changed, text.replace('firstSlideNum="1"', ''));
        assert.deepEqual(current, encode(changed)); assert.equal((await inspectPptx(next)).firstSlideNumber, value ?? 1);
        for (const [name, data] of before) if (name !== 'deck/main.data') assert.deepEqual(after.get(name), data);
        bytes = next;
      } finally { session.dispose(); }
    }
  }
});


test('attribute-v2 browser insertion matches Node and rejects mixed edits and stale plans', async t => {
  const deck = createPresentation(); deck.addSlide({ elements: [{ type: 'text', box: { x: 10, y: 10, width: 100, height: 30 }, content: 'before' }] });
  const bytes = (await generatePptx(deck)).bytes, session = await importPptx(bytes, { projection: 'content', firstSlideNumberSaveProfile: 'attribute-v2' });
  let expected;
  try { session.document.setFirstSlideNumber(1); const plan = await session.prepareSave({ mode: 'edit' }); assert.equal(plan.status, 'ready'); expected = (await session.save(plan.plan)).bytes; } finally { session.dispose(); }
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
  for (const name of ['core', 'pptx-parser']) await page.addScriptTag({ content: await readFile(`packages/${name}/dist/browser/global.js`, 'utf8') });
  const actual = await page.evaluate(async input => { const s = await PPTKit.importPptx(new Uint8Array(input), { projection: 'content', firstSlideNumberSaveProfile: 'attribute-v2' }); try {
    s.document.setFirstSlideNumber(1); const p = await s.prepareSave({ mode: 'edit' }), out = await s.save(p.plan);
    const slide = s.document.slides[0], original = slide.elements[0], changed = structuredClone(original); changed.content[0].runs[0].text = 'mixed'; slide.replaceElement(original.id, changed);
    const mixed = (await s.prepareSave({ mode: 'edit' })).status; slide.replaceElement(original.id, original);
    const fresh = await s.prepareSave({ mode: 'edit' }); s.document.setFirstSlideNumber(2); let stale; try { await s.save(fresh.plan); } catch (e) { stale = e.code; }
    return { bytes: Array.from(out.bytes), mixed, stale };
  } finally { s.dispose(); } }, Array.from(bytes));
  assert.deepEqual(new Uint8Array(actual.bytes), expected); assert.equal(actual.mixed, 'blocked'); assert.equal(actual.stale, 'save-snapshot-stale'); assert.deepEqual(requests, []);
});


test('attribute-v2 with text profile saves slide text and starting number consistently in Node and browser', async t => {
  const deck = createPresentation(); deck.addSlide({ elements: [{ type: 'text', box: { x: 10, y: 10, width: 100, height: 30 }, content: 'before' }] });
  const bytes = (await generatePptx(deck)).bytes, options = { projection: 'content', firstSlideNumberSaveProfile: 'attribute-v2', editSaveProfile: 'text-runs-v3' };
  const session = await importPptx(bytes, options); let expected;
  try {
    const slide = session.document.slides[0], original = slide.elements[0], change = structuredClone(original); change.content[0].runs[0].text = 'after'; slide.replaceElement(original.id, change); session.document.setFirstSlideNumber(17);
    const plan = await session.prepareSave({ mode: 'edit' }); assert.equal(plan.status, 'ready', JSON.stringify(plan.report)); assert.equal(plan.plan.report.changes.length, 2);
    expected = (await session.save(plan.plan)).bytes;
    slide.replaceElement(original.id, original); await assert.rejects(session.save(plan.plan), e => e.code === 'save-snapshot-stale');
  } finally { session.dispose(); }
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
  for (const name of ['core', 'pptx-parser']) await page.addScriptTag({ content: await readFile(`packages/${name}/dist/browser/global.js`, 'utf8') });
  const output = await page.evaluate(async ({ input, options }) => { const s = await PPTKit.importPptx(new Uint8Array(input), options); try {
    const slide = s.document.slides[0], original = slide.elements[0], change = structuredClone(original); change.content[0].runs[0].text = 'after'; slide.replaceElement(original.id, change); s.document.setFirstSlideNumber(17);
    const p = await s.prepareSave({ mode: 'edit' }); return { status: p.status, bytes: Array.from((await s.save(p.plan)).bytes) };
  } finally { s.dispose(); } }, { input: Array.from(bytes), options });
  assert.equal(output.status, 'ready'); assert.deepEqual(new Uint8Array(output.bytes), expected); assert.deepEqual(requests, []);
  const limited = await importPptx(bytes, { ...options, sessionLimits: { maxSaveChanges: 1 } });
  try { const slide = limited.document.slides[0], original = slide.elements[0], change = structuredClone(original); change.content[0].runs[0].text = 'after'; slide.replaceElement(original.id, change); limited.document.setFirstSlideNumber(17); await assert.rejects(limited.prepareSave({ mode: 'edit' }), e => e.code === 'save-limit-exceeded'); } finally { limited.dispose(); }
});


test('all text-run profiles support combined slide-number edits and preserve their own guards', async () => {
  const deck = createPresentation(); deck.addSlide({ elements: [{ type: 'text', box: { x: 10, y: 10, width: 100, height: 30 }, content: 'before' }] });
  const bytes = (await generatePptx(deck)).bytes;
  for (const profile of ['text-runs-v1', 'text-runs-v2', 'text-runs-v3']) {
    const s = await importPptx(bytes, { projection: 'content', firstSlideNumberSaveProfile: 'attribute-v2', editSaveProfile: profile });
    try {
      const slide = s.document.slides[0], original = slide.elements[0], edit = structuredClone(original); edit.content[0].runs[0].text = 'after'; slide.replaceElement(original.id, edit); s.document.setFirstSlideNumber(4);
      const p = await s.prepareSave({ mode: 'edit' }); assert.equal(p.status, 'ready', JSON.stringify(p.report));
      const output = await s.save(p.plan), inspection = await inspectPptx(output.bytes); assert.equal(inspection.firstSlideNumber, 4);
      const bad = structuredClone(edit); bad.box.x += 1; slide.replaceElement(original.id, bad); assert.equal((await s.prepareSave({ mode: 'edit' })).status, 'blocked');
    } finally { s.dispose(); }
  }
});
