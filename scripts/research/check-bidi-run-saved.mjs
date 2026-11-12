import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { importPptx } from '../../packages/pptx-parser/dist/index.js';

const path = process.argv[2], mode = process.argv[3];
assert.ok(path && ['unedited', 'edited', 'both-edited'].includes(mode), 'usage: check-bidi-run-saved.mjs FILE unedited|edited|both-edited');
const bytes = await readFile(path);
const session = await importPptx(bytes, { projection: 'content', textFontProfile: 'script-families-v1', editSaveProfile: 'text-runs-v3' });
try {
  const elements = session.document.slides[0].elements;
  assert.equal(elements.length, 4);
  const expected = new Map([
    ['ltr-control', 'ABC مرحبا - 123 XYZ'],
    ['ltr-run-rtl', `ABC مرحبا - ${mode === 'unedited' ? '123' : '456'} XYZ`],
    ['rtl-control', 'ABC مرحبا - 123 XYZ'],
    ['rtl-run-rtl', `ABC مرحبا - ${mode === 'both-edited' ? '789' : '123'} XYZ`],
  ]);
  for (const element of elements) {
    assert.ok(expected.has(element.name));
    const runs = element.content.flatMap(paragraph => paragraph.runs);
    assert.equal(runs.map(run => run.text).join(''), expected.get(element.name));
    if (element.name.endsWith('run-rtl')) {
      const directed = runs.filter(run => run.style.rightToLeft === true).map(run => run.text).join('');
      assert.ok(directed.includes('مرحبا'));
      const digits = expected.get(element.name).match(/\d+/)[0];
      assert.ok(directed.includes(digits));
    } else assert.ok(runs.every(run => run.style.rightToLeft !== true));
    expected.delete(element.name);
  }
  assert.equal(expected.size, 0);
  const prepared = await session.prepareSave({ mode: 'no-op' });
  assert.equal(prepared.status, 'ready');
  assert.deepEqual(Buffer.from((await session.save(prepared.plan)).bytes), bytes);
  console.log(`${mode}: projected text, run direction and byte-identical no-op verified`);
} finally { session.dispose(); }
