import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { normalizePresentation } from '../../core/dist/index.js';
import { importPptx } from '../dist/index.js';
import { importPptxFile } from '../dist/node.js';
import { encodeStructure } from './structure-fixture.mjs';
import { notesFixture, notesShape } from './notes-fixture.mjs';
import { body } from './content-fixture.mjs';
import { readGeneratedZip } from '../../../tests/support/read-generated-zip.mjs';

const id = '38D9DC4A-2E99-E34A-A7A3-8E01BBC23A4A';
const field = `<a:fld id="{${id}}" type="slidenum"><a:rPr b="1"/><a:t>42</a:t></a:fld>`;
const fixture = (middle, strict = false) => encodeStructure(notesFixture(notesShape({ text: body(`<a:p><a:r><a:t>before</a:t></a:r>${middle}<a:r><a:t>after</a:t></a:r></a:p>`) }), strict));
const content = session => normalizePresentation(session.document).slides.find(slide => slide.notes[0]?.runs.some(run => run.text === 'before')).notes[0].runs;

for (const strict of [false, true]) test(`opt-in notes page-number fields retain cache, identity and adjacent text (${strict ? 'Strict' : 'Transitional'})`, async () => {
  const bytes = fixture(field, strict), session = await importPptx(bytes, { projection: 'content', fieldTextProfile: 'slide-number-v1' });
  try {
    const runs = content(session); assert.deepEqual(runs.map(run => run.text), ['before', '42', 'after']); assert.deepEqual(runs[1].field, { id, kind: 'slide-number' }); assert.equal(runs[1].style.bold, true);
    assert.equal(session.report.fieldTextProfile, 'slide-number-v1');
    const plan = await session.prepareSave({ mode: 'no-op' }); assert.equal(plan.status, 'ready'); assert.deepEqual(Buffer.from((await session.save(plan.plan)).bytes), bytes);
  } finally { session.dispose(); }
  const baseline = await importPptx(bytes, { projection: 'content' });
  try { assert.deepEqual(content(baseline).map(run => run.text), ['before', 'after']); assert.equal(baseline.report.fieldTextProfile, undefined); } finally { baseline.dispose(); }
});

test('unknown or malformed fields remain unmapped and preservation profiles cannot be mixed', async () => {
  for (const invalid of [field.replace('slidenum', 'datetimeFigureOut'), field.replace(id, 'bad'), field.replace('</a:fld>', '<a:t>other</a:t></a:fld>'), field.replace(' type="slidenum"', ' type="slidenum" extra="x"'), field.replace('<a:t>42</a:t>', '<a:t>42<a:r/></a:t>'), field.replace('</a:fld>', 'unexpected</a:fld>')]) {
    const session = await importPptx(fixture(invalid), { projection: 'content', fieldTextProfile: 'slide-number-v1' });
    try { assert.deepEqual(content(session).map(run => run.text), ['before', 'after']); assert.ok(session.report.findings.some(f => f.feature.includes('unmapped-property'))); } finally { session.dispose(); }
  }
  for (const options of [{ projection: 'structure', fieldTextProfile: 'slide-number-v1' }, { projection: 'content', fieldTextProfile: 'invalid' }, { projection: 'content', fieldTextProfile: 'slide-number-v1', pictureEditSaveProfile: 'replace-v1' }]) await assert.rejects(importPptx(fixture(field), options), TypeError);
});

test('Node file import snapshots the opt-in profile before asynchronous file access', async t => {
  const directory = await mkdtemp(path.join(tmpdir(), 'pptkit-field-')); t.after(() => rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'source.pptx'); await writeFile(file, fixture(field));
  const options = { projection: 'content', fieldTextProfile: 'slide-number-v1' }, pending = importPptxFile(file, options); options.fieldTextProfile = undefined;
  const imported = await pending;
  try { assert.equal(imported.session.report.fieldTextProfile, 'slide-number-v1'); assert.equal(content(imported.session)[1].field.kind, 'slide-number'); } finally { imported.session.dispose(); }
});

for (const strict of [false, true]) test(`adjacent notes text saves while field cache and identity remain blocked (${strict ? 'Strict' : 'Transitional'})`, async () => {
  let bytes = fixture(field, strict);
  for (let round = 0; round < 3; round++) {
    const session = await importPptx(bytes, { projection: 'content', fieldTextProfile: 'slide-number-v1', editSaveProfile: 'text-runs-v2', ...(round === 1 ? { notesTextProfile: 'master-style-v1', textFontProfile: 'script-families-v3' } : {}) });
    try {
      const slide = session.document.slides.find(item => Array.isArray(item.notes) && item.notes[0]?.runs.some(run => run.field)), original = structuredClone(slide.notes), runs = original[0].runs;
      const fieldBinding = session.report.notesTextRuns.find(binding => binding.slideId === slide.id && binding.runIndex === 1);
      assert.equal(fieldBinding.status, 'unmapped'); assert.match(fieldBinding.reason, /field caches/);
      const modifiedCache = structuredClone(original); modifiedCache[0].runs[1].text = '999'; slide.setNotes(modifiedCache); assert.equal((await session.prepareSave({ mode: 'edit' })).status, 'blocked');
      modifiedCache[0].runs[0].text = 'also changed'; slide.setNotes(modifiedCache); assert.equal((await session.prepareSave({ mode: 'edit' })).status, 'blocked');
      const modifiedId = structuredClone(original); modifiedId[0].runs[1].field.id = '11111111-2222-3333-4444-555555555555'; slide.setNotes(modifiedId); assert.equal((await session.prepareSave({ mode: 'edit' })).status, 'blocked');
      runs[0].text += '!'; slide.setNotes(original); const plan = await session.prepareSave({ mode: 'edit' }); assert.equal(plan.status, 'ready', JSON.stringify(plan.report));
      const saved = Buffer.from((await session.save(plan.plan)).bytes), before = readGeneratedZip(bytes), after = readGeneratedZip(saved);
      for (const [name, payload] of before) {
        if (name === 'deck/notes/z.xml') { const originalXml = payload.toString(), nextXml = after.get(name).toString(); assert.equal(nextXml, originalXml.replace(`before${'!'.repeat(round)}</a:t>`, `before${'!'.repeat(round + 1)}</a:t>`)); assert.ok(nextXml.includes(field)); }
        else assert.deepEqual(after.get(name), payload);
      }
      bytes = saved;
    } finally { session.dispose(); }
  }
});

for (const strict of [false, true]) test(`explicit date profile preserves supported native format, cache and adjacent edits (${strict ? 'Strict' : 'Transitional'})`, async () => {
  const types = [['datetime', 'application-default'], ['datetime1', 'month-day-year'], ['datetime2', 'weekday-month-day-year'], ['datetime3', 'day-month-year'], ['datetime4', 'month-day-year-long'], ['datetimeFigureOut', 'automatic-month-day-year']];
  for (const [type, updateFormat] of types) {
    const dateField = field.replace('type="slidenum"', `type="${type}"`); let bytes = fixture(dateField, strict);
    const legacy = await importPptx(bytes, { projection: 'content', fieldTextProfile: 'slide-number-v1' });
    try { assert.deepEqual(content(legacy).map(run => run.text), ['before', 'after']); } finally { legacy.dispose(); }
    for (let round = 0; round < 3; round++) {
      const s = await importPptx(bytes, { projection: 'content', fieldTextProfile: 'date-and-number-v1', editSaveProfile: 'text-runs-v2' });
      try {
        const slide = s.document.slides.find(item => Array.isArray(item.notes) && item.notes[0]?.runs.some(run => run.field)), notes = structuredClone(slide.notes), entry = notes[0].runs[1];
        assert.deepEqual(entry.field, { id, kind: 'dynamic-date', updateFormat }); assert.equal(entry.text, '42'); assert.equal(s.report.fieldTextProfile, 'date-and-number-v1');
        const bad = structuredClone(notes); bad[0].runs[1].field.updateFormat = 'application-default' === updateFormat ? 'day-month-year' : 'application-default'; slide.setNotes(bad); assert.equal((await s.prepareSave({ mode: 'edit' })).status, 'blocked');
        notes[0].runs[0].text += '!'; slide.setNotes(notes); const plan = await s.prepareSave({ mode: 'edit' }); assert.equal(plan.status, 'ready', JSON.stringify(plan.report));
        const result = Buffer.from((await s.save(plan.plan)).bytes), before = readGeneratedZip(bytes), after = readGeneratedZip(result);
        for (const [name, payload] of before) { if (name === 'deck/notes/z.xml') { const a = payload.toString(), b = after.get(name).toString(); assert.equal(b, a.replace(`before${'!'.repeat(round)}</a:t>`, `before${'!'.repeat(round + 1)}</a:t>`)); assert.ok(b.includes(dateField)); } else assert.deepEqual(after.get(name), payload); }
        bytes = result;
      } finally { s.dispose(); }
    }
  }
});

test('date profile does not infer unsupported native types or malformed fields', async () => {
  for (const type of ['datetime5', 'datetime13', 'DATETIME1', 'constructor']) {
    const s = await importPptx(fixture(field.replace('slidenum', type)), { projection: 'content', fieldTextProfile: 'date-and-number-v1' });
    try { assert.deepEqual(content(s).map(run => run.text), ['before', 'after']); assert.ok(s.report.findings.some(finding => finding.feature.includes('unmapped-property'))); } finally { s.dispose(); }
  }
});

for (const strict of [false, true]) test(`date-time profile preserves all remaining formats and protected fields (${strict ? 'Strict' : 'Transitional'})`, async () => {
  const types = [['datetime5', 'day-short-month-year'], ['datetime6', 'month-year'], ['datetime7', 'short-month-year'], ['datetime8', 'month-day-year-time-12'], ['datetime9', 'month-day-year-time-12-seconds'], ['datetime10', 'time-24'], ['datetime11', 'time-24-seconds'], ['datetime12', 'time-12'], ['datetime13', 'time-12-seconds']];
  for (const [type, updateFormat] of types) {
    const dateField = field.replace('type="slidenum"', `type="${type}"`); let bytes = fixture(dateField, strict);
    const legacy = await importPptx(bytes, { projection: 'content', fieldTextProfile: 'date-and-number-v1' });
    try { assert.deepEqual(content(legacy).map(run => run.text), ['before', 'after']); } finally { legacy.dispose(); }
    for (let round = 0; round < 3; round++) {
      const s = await importPptx(bytes, { projection: 'content', fieldTextProfile: 'date-time-and-number-v1', editSaveProfile: 'text-runs-v2' });
      try {
        const slide = s.document.slides.find(item => Array.isArray(item.notes) && item.notes[0]?.runs.some(run => run.field)), notes = structuredClone(slide.notes);
        assert.deepEqual(notes[0].runs[1].field, { id, kind: 'dynamic-date', updateFormat }); assert.equal(notes[0].runs[1].text, '42');
        const bad = structuredClone(notes); bad[0].runs[1].text = '999'; slide.setNotes(bad); assert.equal((await s.prepareSave({ mode: 'edit' })).status, 'blocked');
        notes[0].runs[0].text += '!'; slide.setNotes(notes); const plan = await s.prepareSave({ mode: 'edit' }); assert.equal(plan.status, 'ready', JSON.stringify(plan.report));
        const result = Buffer.from((await s.save(plan.plan)).bytes), before = readGeneratedZip(bytes), after = readGeneratedZip(result);
        for (const [name, payload] of before) { if (name === 'deck/notes/z.xml') { const a = payload.toString(), b = after.get(name).toString(); assert.equal(b, a.replace(`before${'!'.repeat(round)}</a:t>`, `before${'!'.repeat(round + 1)}</a:t>`)); assert.ok(b.includes(dateField)); } else assert.deepEqual(after.get(name), payload); }
        bytes = result;
      } finally { s.dispose(); }
    }
  }
});


for (const strict of [false, true]) test(`page-number attribute and adjacent notes text save atomically across three rounds (${strict ? 'Strict' : 'Transitional'})`, async () => {
  let bytes = fixture(field, strict), previous;
  for (const [round, value] of [17, 0, undefined].entries()) {
    const session = await importPptx(bytes, { projection: 'content', fieldTextProfile: 'slide-number-v1', editSaveProfile: 'text-runs-v2', firstSlideNumberSaveProfile: 'attribute-v2', notesTextProfile: 'master-style-v1', textFontProfile: 'script-families-v3' });
    try {
      const slide = session.document.slides.find(item => item.notes[0]?.runs.some(run => run.field)), original = structuredClone(slide.notes), notes = structuredClone(original);
      assert.equal(session.document.firstSlideNumber, previous); session.document.setFirstSlideNumber(value);
      notes[0].runs[0].text += '!'; slide.setNotes(notes);
      const illegal = structuredClone(notes); illegal[0].runs[1].text = 'tampered'; slide.setNotes(illegal);
      assert.equal((await session.prepareSave({ mode: 'edit' })).status, 'blocked'); slide.setNotes(notes);
      const result = await session.prepareSave({ mode: 'edit' }); assert.equal(result.status, 'ready', JSON.stringify(result.report));
      assert.equal(result.plan.report.changes.length, 2);
      const saved = Buffer.from((await session.save(result.plan)).bytes), before = readGeneratedZip(bytes), after = readGeneratedZip(saved);
      for (const [name, payload] of before) {
        if (name === 'deck/notes/z.xml') { const xml = after.get(name).toString(); assert.equal(xml, payload.toString().replace(`before${'!'.repeat(round)}</a:t>`, `before${'!'.repeat(round + 1)}</a:t>`)); assert.ok(xml.includes(field)); }
        else if (name === 'deck/main.data') { const old = payload.toString(), current = after.get(name).toString();
          if (previous === undefined) assert.equal(current.replace(` firstSlideNum="${value}"`, ''), old);
          else if (value === undefined) assert.equal(current, old.replace(`firstSlideNum="${previous}"`, ''));
          else assert.equal(current, old.replace(`firstSlideNum="${previous}"`, `firstSlideNum="${value}"`));
        } else assert.deepEqual(after.get(name), payload);
      }
      previous = value; bytes = saved;
    } finally { session.dispose(); }
  }
});
