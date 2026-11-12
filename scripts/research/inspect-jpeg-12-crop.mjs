import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { importPptx } from '../../packages/pptx-parser/dist/index.js';

const bytes = await readFile('.tmp/native-jpeg-12/crop.pptx');
const session = await importPptx(bytes, { projection: 'content', pictureEditSaveProfile: 'replace-v1' });
try {
  const pictures = session.document.slides[0].elements.filter(element => element.type === 'image');
  assert.equal(pictures.length, 3);
  assert.deepEqual(pictures.slice(0, 2).map(picture => picture.box.x), [30, 330]);
  assert.equal(pictures[2].fit, 'crop');
  assert.deepEqual(pictures[2].crop, { left: 0.25585, top: 0, right: 0, bottom: 0 });
  assert.ok(Math.abs(pictures[2].box.x - 8845826 / 12700) < 1e-8);
  assert.ok(Math.abs(pictures[2].box.width - 2457174 / 12700) < 1e-8);
  const prepared = await session.prepareSave({ mode: 'no-op' });
  assert.equal(prepared.status, 'ready', JSON.stringify(prepared.report));
  assert.deepEqual(Buffer.from((await session.save(prepared.plan)).bytes), bytes);
  console.log('Imported crop and byte-identical no-op verified.');
} finally { session.dispose(); }
