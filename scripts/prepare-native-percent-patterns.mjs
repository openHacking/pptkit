import { mkdir, writeFile } from 'node:fs/promises';
import { createPresentation } from '../packages/core/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';

const values = [5, 10, 20, 25, 30, 40, 50, 60, 70, 75, 80, 90];
const deck = createPresentation({ size: { width: 960, height: 540 } });
const slide = deck.addSlide({ id: 'patterns' });
for (const [i, value] of values.entries()) {
  slide.addElement({ id: `percent${value}`, name: `percent${value}`, type: 'shape', shape: 'rect', box: { x: 40 + (i % 4) * 230, y: 35 + Math.floor(i / 4) * 165, width: 200, height: 135 }, style: { fill: { type: 'pattern', pattern: `percent${value}`, foreground: '000000', background: 'FFFFFF' }, stroke: { paint: { type: 'solid', color: '808080' }, width: 1 } } });
}
await mkdir('.tmp/native-percent-patterns', { recursive: true });
// Do not overwrite a retained observation.
await writeFile('.tmp/native-percent-patterns/percent-patterns.pptx', (await generatePptx(deck)).bytes, { flag: 'wx' });
