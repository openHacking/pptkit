import { mkdir, writeFile } from 'node:fs/promises';
import { createPresentation } from '../../packages/core/dist/index.js';
import { generatePptx } from '../../packages/pptx-exporter/dist/index.js';

const folder = process.argv[2] ?? '.tmp/vertical-writing';
await mkdir(folder, { recursive: true });
const deck = createPresentation();
const modes = ['horz', 'vert', 'vert270', 'eaVert', 'mongolianVert', 'wordArtVert', 'wordArtVertRtl'];
const content = [{ runs: [{ text: '日本語（ABC 123）', style: { fontSize: 24, fontFamily: 'Arial', eastAsianFontFamily: 'Hiragino Sans' } }] }];
for (const mode of modes) {
  const slide = deck.addSlide();
  slide.addElement({ type: 'text', box: { x: 60, y: 30, width: 700, height: 40 }, content: 'MODE: ' + mode });
  slide.addElement({ type: 'text', box: { x: 60, y: 95, width: 320, height: 315 }, content });
  slide.addElement({ type: 'table', box: { x: 450, y: 95, width: 380, height: 315 }, columns: [190, 190],
    rows: [{ height: 315, cells: [{ content }, { content }] }] });
}
await writeFile(folder + '/horizontal.pptx', (await generatePptx(deck)).bytes);
