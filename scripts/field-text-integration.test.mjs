import test from 'node:test';
import assert from 'node:assert/strict';
import { createPresentation, materializeFieldText, normalizePresentation } from '../packages/core/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';
import { renderPresentationToSvg } from '../packages/svg-renderer/dist/index.js';
import { importPptx } from '../packages/pptx-parser/dist/index.js';
import { readGeneratedZip } from '../tests/support/read-generated-zip.mjs';

test('explicit static field snapshot becomes visible authored text in SVG and PPTX', async () => {
  const fields = [{ runs: [{ text: 'Page ' }, { field: { id: 'number', kind: 'slide-number' }, style: { bold: true } }, { text: ' — ' }, { field: { id: 'footer', kind: 'footer', text: 'Review' } }] }];
  const snapshot = materializeFieldText(fields, { pageOrder: ['one', 'two'], pageId: 'two', firstPageNumber: 1 });
  const deck = createPresentation(); deck.addSlide({ id: 'one', notes: snapshot.content, elements: [{ type: 'text', content: snapshot.content, box: { x: 30, y: 30, width: 400, height: 60 } }, { type: 'table', box: { x: 30, y: 120, width: 400, height: 60 }, columns: [400], rows: [{ height: 60, cells: [{ content: snapshot.content }] }] }] });
  const svg = (await renderPresentationToSvg(deck)).slides[0].svg;
  assert.match(svg, />Page </); assert.match(svg, />2</); assert.match(svg, />Review</);
  const bytes = (await generatePptx(deck)).bytes, parts = readGeneratedZip(bytes), xml = parts.get('ppt/slides/slide1.xml').toString(), notes = parts.get('ppt/notesSlides/notesSlide1.xml').toString();
  assert.match(xml, /<a:t>Page <\/a:t>/); assert.match(xml, /<a:t>2<\/a:t>/); assert.match(xml, /<a:t>Review<\/a:t>/); assert.doesNotMatch(xml, /<a:fld\b/);
  assert.equal((xml.match(/<a:t>Review<\/a:t>/g) ?? []).length, 2); assert.match(notes, /<a:t>Review<\/a:t>/); assert.doesNotMatch(notes, /<a:fld\b/);
  assert.deepEqual(snapshot.fields.map(f => f.value), ['2', 'Review']);
  let current = bytes;
  for (let round = 0; round < 3; round++) {
    const session = await importPptx(current, { projection: 'content' });
    try {
      const slide = normalizePresentation(session.document).slides[0], text = slide.elements.find(element => element.type === 'text'), table = slide.elements.find(element => element.type === 'table');
      assert.equal(text.plainText, 'Page 2 — Review');
      assert.equal(table.rows[0].cells[0].content[0].runs.map(run => run.text).join(''), 'Page 2 — Review');
      assert.equal(slide.notes[0].runs.map(run => run.text).join(''), 'Page 2 — Review');
      const prepared = await session.prepareSave({ mode: 'no-op' }); assert.equal(prepared.status, 'ready');
      const next = (await session.save(prepared.plan)).bytes; assert.deepEqual(next, current); current = next;
    } finally { session.dispose(); }
  }
});
