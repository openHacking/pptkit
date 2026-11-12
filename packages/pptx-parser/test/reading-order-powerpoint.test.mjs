import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { readFile } from 'node:fs/promises';
import test from 'node:test';
import { importPptx } from '../dist/index.js';

test('PowerPoint reading order resave projects changed drawing order without inventing independent metadata', async () => {
  const bytes = new Uint8Array(await readFile(new URL('../../../tests/fixtures/reading-order-powerpoint-resave.pptx', import.meta.url)));
  const evidence = JSON.parse(await readFile(new URL('../../../docs/testing/powerpoint-reading-order-evidence.json', import.meta.url), 'utf8'));
  assert.equal(createHash('sha256').update(bytes).digest('hex'), evidence.saved.sha256);
  const session = await importPptx(bytes, { projection: 'content' });
  try {
    const slide = session.document.slides[0];
    assert.equal(slide.readingOrder, undefined);
    assert.deepEqual(slide.elements.map(element => [element.name, element.accessibility.description, element.box.x, element.box.y]), [
      ['Blue rectangle', 'Blue rectangle', 245, 140],
      ['Red rectangle', 'Red rectangle', 90, 65],
    ]);
    const candidate = await session.prepareSave({ mode: 'no-op' });
    assert.equal(candidate.status, 'ready');
    assert.deepEqual((await session.save(candidate.plan)).bytes, bytes);
  } finally { session.dispose(); }
});
