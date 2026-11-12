import { mkdir, writeFile } from 'node:fs/promises';
import { createPresentation } from '../../packages/core/dist/index.js';
import { generatePptx } from '../../packages/pptx-exporter/dist/index.js';

const folder = process.argv[2] ?? '.tmp/vertical-writing';
await mkdir(folder, { recursive: true });
const deck = createPresentation(), slide = deck.addSlide();
const textDirection = 'eastAsianVertical';
const content = [{ style: { lineSpacing: 1.5 }, runs: [{ text: '中文ABCD', style: {
  fontFamily: 'Arial', eastAsianFontFamily: 'Hiragino Sans GB', fontSize: 20, language: 'zh-CN',
} }] }];
for (const [x, height, owner] of [[60, 180, 'FRAME'], [260, 55, 'FRAME'], [460, 180, 'CELL'], [660, 55, 'CELL']]) {
  slide.addElement({ type: 'text', box: { x, y: 30, width: 155, height: 35 }, content: owner + ' ' + height });
  const box = { x, y: 110, width: 150, height };
  if (owner === 'FRAME') slide.addElement({ type: 'text', box, content, frame: { textDirection, margin: 5 } });
  else slide.addElement({ type: 'table', box, columns: [150], rows: [{ height, cells: [{ content, style: { textDirection, margin: 5 } }] }] });
}
await writeFile(folder + '/east-asian-wrap-control.pptx', (await generatePptx(deck)).bytes, { flag: 'wx' });
