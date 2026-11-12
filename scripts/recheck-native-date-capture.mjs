// Read-only verification of the named optional LibreOffice capture; no native app is launched.
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { readFile } from 'node:fs/promises';
import { normalizePresentation } from '../packages/core/dist/index.js';
import { importPptx } from '../packages/pptx-parser/dist/index.js';

const evidence = JSON.parse(await readFile('docs/testing/native-date-libreoffice-evidence.json', 'utf8'));
assert.equal(evidence.files.length, 4); assert.equal(evidence.records.length, 4);
for (const [index, file] of evidence.files.entries()) {
  const bytes = await readFile(file.file); assert.equal(createHash('sha256').update(bytes).digest('hex'), file.sha256);
  const session = await importPptx(bytes, { projection: 'content', fieldTextProfile: 'date-and-number-v1' });
  try {
    const doc = normalizePresentation(session.document), runs = doc.slides[0].elements.flatMap(element => (element.type === 'text' ? element.content : element.type === 'shape' ? element.text?.content : [])?.flatMap(paragraph => paragraph.runs) ?? []), field = runs.filter(run => run.field);
    assert.equal(field.length, 1); assert.equal(field[0].field.kind, 'dynamic-date'); assert.equal(field[0].field.updateFormat, 'day-month-year');
    assert.equal(`{${field[0].field.id.toUpperCase()}}`, evidence.records[index].fields[0].id); assert.equal(field[0].text, evidence.records[index].fields[0].cache);
    const plan = await session.prepareSave({ mode: 'no-op' }); assert.equal(plan.status, 'ready'); assert.deepEqual(Buffer.from((await session.save(plan.plan)).bytes), bytes);
  } finally { session.dispose(); }
}
console.log(JSON.stringify({ capturedInputs: evidence.files.length, exactNoOps: evidence.files.length, producer: evidence.application }));
