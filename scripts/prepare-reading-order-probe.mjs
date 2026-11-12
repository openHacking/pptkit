import { writeFile } from 'node:fs/promises';
import { createPresentation } from '../packages/core/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';

const destination = process.argv[2];
if (!destination) throw new TypeError('Provide a new path for the probe PPTX.');
const deck = createPresentation({ size: { width: 640, height: 360 } });
const shape = (id, name, color, x, y) => ({
  id, name, type: 'shape', shape: 'rect', box: { x, y, width: 270, height: 170 },
  accessibility: { description: name }, style: { fill: { type: 'solid', color }, stroke: { paint: { type: 'none' } } },
});
deck.addSlide({
  id: 'overlap',
  elements: [shape('red', 'Red rectangle', 'D02020', 90, 65), shape('blue', 'Blue rectangle', '2040D0', 245, 140)],
  readingOrder: [{ elementIds: ['blue', 'red'] }],
});
const result = await generatePptx(deck);
if (!result.warnings.some(warning => warning.code === 'reading-order-not-exported')) throw new Error('Missing independent reading order omission warning.');
await writeFile(destination, result.bytes, { flag: 'wx' });
