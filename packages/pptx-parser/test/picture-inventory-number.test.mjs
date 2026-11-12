import test from 'node:test';
import assert from 'node:assert/strict';
import { importPptx } from '../dist/index.js';
import { importPptxFile } from '../dist/node.js';
import { mkdtemp, writeFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { imageFixture, encodeImages, picture } from './image-fixture.mjs';
import { shape, body } from './content-fixture.mjs';
import { readGeneratedZip } from '../../../tests/support/read-generated-zip.mjs';

const field = '<a:fld id="{38D9DC4A-2E99-E34A-A7A3-8E01BBC23A4A}" type="slidenum"><a:t>42</a:t></a:fld>';
const options = { projection: 'content', pictureSourceProfile: 'embedded-v1', firstSlideNumberSaveProfile: 'attribute-v2', editSaveProfile: 'text-runs-v3', fieldTextProfile: 'slide-number-v1' };

for (const strict of [false, true]) test(`picture inventory survives combined number and text saves without granting image edits (${strict ? 'Strict' : 'Transitional'})`, async () => {
  const content = picture({ id: 2 }) + picture({ id: 3 }) + shape({ id: '20', txBox: true, text: body(`<a:p><a:r><a:t>before</a:t></a:r>${field}</a:p>`) });
  let bytes = encodeImages(imageFixture(content, strict)), previous;
  for (const [round, value] of [17, 0, undefined].entries()) {
    const s = await importPptx(bytes, options);
    try {
      assert.equal(s.document.firstSlideNumber, previous); assert.equal(s.report.pictureSources.length, 2);
      const [a, b] = s.report.pictureSources; assert.equal(a.mediaPart, b.mediaPart); assert.notEqual(a.sourceElementIndex, b.sourceElementIndex);
      const slide = s.document.slides[0], original = slide.elements[2], changed = structuredClone(original); changed.content[0].runs[0].text += '!'; slide.replaceElement(original.id, changed); s.document.setFirstSlideNumber(value);
      const pic = slide.elements[0], modified = structuredClone(pic); modified.box.x++; slide.replaceElement(pic.id, modified); assert.equal((await s.prepareSave({ mode: 'edit' })).status, 'blocked'); slide.replaceElement(pic.id, pic);
      const p = await s.prepareSave({ mode: 'edit' }); assert.equal(p.status, 'ready', JSON.stringify(p.report)); assert.equal(p.plan.report.changes.length, 2);
      const next = (await s.save(p.plan)).bytes, before = readGeneratedZip(bytes), after = readGeneratedZip(next);
      for (const [name, payload] of before) {
        if (name === 'deck/z.xml') { const source = payload.toString(); assert.equal(after.get(name).toString(), source.replace(`before${'!'.repeat(round)}</a:t>`, `before${'!'.repeat(round + 1)}</a:t>`)); assert.ok(after.get(name).toString().includes(field)); }
        else if (name !== 'deck/main.data') assert.deepEqual(after.get(name), payload, name);
      }
      const reopened = await importPptx(next, options); try { assert.equal(reopened.document.firstSlideNumber, value); assert.deepEqual(reopened.report.pictureSources, s.report.pictureSources); } finally { reopened.dispose(); }
      previous = value; bytes = next;
    } finally { s.dispose(); }
  }
});


test('file import captures both options and stale/limited plans remain atomic', async t => {
  const directory = await mkdtemp(path.join(tmpdir(), 'pptkit-inventory-number-')); t.after(() => rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'source.pptx'), bytes = encodeImages(imageFixture()); await writeFile(file, bytes);
  const config = { projection: 'content', pictureSourceProfile: 'embedded-v1', firstSlideNumberSaveProfile: 'attribute-v2' }, pending = importPptxFile(file, config);
  config.pictureSourceProfile = undefined; config.firstSlideNumberSaveProfile = undefined;
  const { session } = await pending;
  try {
    assert.equal(session.report.pictureSources.length, 1); session.document.setFirstSlideNumber(17); const p = await session.prepareSave({ mode: 'edit' }); assert.equal(p.status, 'ready');
    session.document.setFirstSlideNumber(18); await assert.rejects(session.save(p.plan), e => e.code === 'save-snapshot-stale');
    const cancelled = new AbortController(); cancelled.abort(); await assert.rejects(session.prepareSave({ mode: 'edit', signal: cancelled.signal }), e => e.code === 'operation-aborted');
    session.document.setFirstSlideNumber(17); assert.deepEqual((await session.save(p.plan)).bytes, (await session.save((await session.prepareSave({ mode: 'edit' })).plan)).bytes);
  } finally { session.dispose(); }
  await assert.rejects(importPptx(bytes, { projection: 'content', pictureSourceProfile: 'embedded-v1', firstSlideNumberSaveProfile: 'existing-v1' }), TypeError);
});
