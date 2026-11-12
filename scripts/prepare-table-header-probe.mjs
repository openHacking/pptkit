import { writeFile } from 'node:fs/promises';
import { createPresentation } from '../packages/core/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';

const destination = process.argv[2];
if (!destination) throw new TypeError('Provide a new path for the probe PPTX.');
const deck = createPresentation({ size: { width: 640, height: 360 } });
deck.addSlide({ elements: [{
  type: 'table', name: 'Header experiment', box: { x: 100, y: 80, width: 440, height: 170 },
  columns: [220, 220],
  rows: [['Region', 'Total'], ['North', '25']].map(values => ({ cells: values.map(content => ({ content })) })),
}] });
await writeFile(destination, (await generatePptx(deck)).bytes, { flag: 'wx' });
