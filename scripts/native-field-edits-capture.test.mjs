import test from 'node:test';
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { readFile } from 'node:fs/promises';
import { importPptx } from '../packages/pptx-parser/dist/index.js';
import { normalizePresentation } from '../packages/core/dist/index.js';

test('captured PowerPoint reorder and adjacent text edit reimport and preserve exact source', async () => {
  const evidence = JSON.parse(await readFile('docs/testing/native-field-edits-powerpoint-evidence.json', 'utf8'));
  for (const record of evidence.records) {
    const bytes = await readFile(record.file);
    assert.equal(createHash('sha256').update(bytes).digest('hex'), record.sha256);
    const session = await importPptx(bytes, { projection: 'content', fieldTextProfile: 'date-time-and-number-v1' });
    try {
      const document = normalizePresentation(session.document);
      assert.equal(document.slides.length, 2);
      for (let i = 0; i < 2; i++) {
        const element = document.slides[i].elements.find(element => element.type === 'text');
        assert.ok(element);
        const runs = element.content.flatMap(paragraph => paragraph.runs);
        assert.equal(runs.map(run => run.text).join(''), record.slides[i].text);
        const fields = runs.filter(run => run.field).map(run => {
          if (run.field.kind !== 'slide-number') { assert.equal(run.field.kind, 'dynamic-date'); assert.equal(run.field.updateFormat, 'day-month-year'); }
          return { id: `{${run.field.id.toUpperCase()}}`, type: run.field.kind === 'slide-number' ? 'slidenum' : 'datetime3', cache: run.text };
        });
        assert.deepEqual(fields, record.slides[i].fields);
      }
      const plan = await session.prepareSave({ mode: 'no-op' });
      assert.equal(plan.status, 'ready');
      assert.deepEqual(Buffer.from((await session.save(plan.plan)).bytes), bytes);
    } finally { session.dispose(); }
  }
});
