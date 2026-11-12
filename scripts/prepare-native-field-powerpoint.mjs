import { writeFile } from 'node:fs/promises';
import { createPresentation } from '../packages/core/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';

const deck = createPresentation();
for (let i = 1; i <= 2; i++) {
  const slide = deck.addSlide({ id: `page-${i}` });
  slide.addElement({ id: `field-${i}`, type: 'text', box: { x: 60, y: 70, width: 700, height: 100 }, content: [{ runs: [
    { text: `Page ${i}: ` },
    { text: 'STALE-NUMBER', field: { id: `11111111-2222-3333-4444-${String(i).padStart(12, '0')}`, kind: 'slide-number' } },
    { text: ' / Date: ' },
    { text: 'STALE-DATE', field: { id: `aaaaaaaa-bbbb-cccc-dddd-${String(i).padStart(12, '0')}`, kind: 'dynamic-date', updateFormat: 'day-month-year' } },
  ] }] });
}
await writeFile('.tmp/native-field-powerpoint-01/source.pptx', (await generatePptx(deck)).bytes, { flag: 'wx' });
