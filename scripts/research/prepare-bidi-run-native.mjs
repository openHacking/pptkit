import { createPresentation } from '../../packages/core/dist/index.js';
import { generatePptx } from '../../packages/pptx-exporter/dist/index.js';
import { mkdir, writeFile } from 'node:fs/promises';
const dir = process.argv[2] ?? '.tmp/bidi-run-native'; const authored = process.argv.includes('--authored'); const fontProbe = process.argv.includes('--font-probe'); await mkdir(dir, { recursive: true });
const deck = createPresentation(), slide = deck.addSlide();
for (const [row, direction] of ['ltr', 'rtl'].entries()) for (const [col, kind] of ['control', 'run-rtl'].entries()) {
  slide.addElement({ type: 'text', id: `${direction}-${kind}`, box: { x: 40 + col * 410, y: 50 + row * 170, width: 350, height: 100 }, frame: { margin: 0, wrap: false },
    content: [{ style: { direction, align: 'left' }, runs: ['ABC ', 'مرحبا', ' - ', '123', ' XYZ'].map((text, index) => ({ text, style: { fontFamily: 'Arial', complexScriptFontFamily: fontProbe ? 'Times New Roman' : 'Arial', fontSize: 32, ...(authored && kind === 'run-rtl' && index >= 1 && index <= 3 ? { rightToLeft: true } : {}) } })) }] });
}
await writeFile(`${dir}/base.pptx`, (await generatePptx(deck)).bytes);
