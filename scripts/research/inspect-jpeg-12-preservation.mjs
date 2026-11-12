import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { importPptx } from '../../packages/pptx-parser/dist/index.js';

for (const [name, x, y] of [['preserved-before.pptx', 8001000, 1651000], ['preserved-native1.pptx', 7941366, 1651000], ['preserved-native2.pptx', 7941366, 1710634], ['preserved-native3.pptx', 8001000, 1710634]]) {
  const bytes = await readFile(`.tmp/native-jpeg-12/${name}`);
  const session = await importPptx(bytes, { projection: 'content', pictureEditSaveProfile: 'replace-v1' });
  try {
    const pictures = session.document.slides[0].elements.filter(element => element.type === 'image');
    assert.equal(pictures.length, 3);
    assert.ok(Math.abs(pictures[2].box.x - x / 12700) < 1e-8);
    assert.ok(Math.abs(pictures[2].box.y - y / 12700) < 1e-8);
    const prepared = await session.prepareSave({ mode: 'no-op' });
    assert.equal(prepared.status, 'ready', JSON.stringify(prepared.report));
    assert.deepEqual(Buffer.from((await session.save(prepared.plan)).bytes), bytes);
  } finally { session.dispose(); }
}
console.log('Candidate and native save: expected image position and byte-identical no-op verified.');
