import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { importPptx } from '../../packages/pptx-parser/dist/index.js';

const bytes = await readFile('.tmp/native-jpeg-12/edit.pptx');
const session = await importPptx(bytes, { projection: 'content', pictureEditSaveProfile: 'replace-v1' });
try {
  const pictures = session.document.slides[0].elements.filter(element => element.type === 'image');
  assert.equal(pictures.length, 3);
  assert.ok(Math.abs(pictures[0].box.x - 440634 / 12700) < 1e-8);
  assert.deepEqual(pictures.slice(1).map(picture => picture.box.x), [330, 630]);
  const prepared = await session.prepareSave({ mode: 'no-op' });
  assert.equal(prepared.status, 'ready', JSON.stringify(prepared.report));
  assert.deepEqual(Buffer.from((await session.save(prepared.plan)).bytes), bytes);
  console.log('Three imported pictures, expected positions and byte-identical no-op verified.');
} finally { session.dispose(); }
