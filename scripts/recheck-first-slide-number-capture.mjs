// Read-only verification of captured files; this does not repeat the application experiment.
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { readFile } from 'node:fs/promises';
import { normalizePresentation } from '../packages/core/dist/index.js';
import { importPptx, inspectPptx } from '../packages/pptx-parser/dist/index.js';

const evidence = JSON.parse(await readFile('docs/testing/first-slide-number-libreoffice-evidence.json', 'utf8'));
assert.equal(evidence.files.length, 2); assert.equal(evidence.records.length, 2);
for (const [index, file] of evidence.files.entries()) {
  const bytes = await readFile(file.file); assert.equal(createHash('sha256').update(bytes).digest('hex'), file.sha256);
  const inspection = await inspectPptx(bytes), recorded = evidence.records[index].firstSlideNum;
  assert.equal(inspection.firstSlideNumberDeclared, recorded !== null);
  assert.equal(inspection.firstSlideNumber, recorded === null ? 1 : Number(recorded));
  const session = await importPptx(bytes, { projection: 'content', fieldTextProfile: 'slide-number-v1' });
  try {
    const doc = normalizePresentation(session.document); assert.equal(doc.slides.length, 2);
    for (let page = 0; page < 2; page++) {
      const fields = doc.slides[page].elements.flatMap(element => (element.type === 'text' ? element.content : element.type === 'shape' ? element.text?.content : [])?.flatMap(paragraph => paragraph.runs) ?? []).filter(run => run.field);
      assert.equal(fields.length, 1); assert.equal(fields[0].field.kind, 'slide-number');
      assert.equal(`{${fields[0].field.id.toUpperCase()}}`, evidence.records[index].fields[page][0].id);
      assert.equal(fields[0].text, evidence.records[index].fields[page][0].cache);
    }
    const plan = await session.prepareSave({ mode: 'no-op' }); assert.equal(plan.status, 'ready'); assert.deepEqual(Buffer.from((await session.save(plan.plan)).bytes), bytes);
  } finally { session.dispose(); }
}
console.log(JSON.stringify({ capturedInputs: 2, exactNoOps: 2, producer: evidence.application }));
