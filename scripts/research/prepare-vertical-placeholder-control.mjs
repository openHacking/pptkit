import { mkdir, writeFile } from 'node:fs/promises';
import { createPresentation } from '../../packages/core/dist/index.js';
import { generatePptx } from '../../packages/pptx-exporter/dist/index.js';

const folder = process.argv[2] ?? '.tmp/vertical-writing';
await mkdir(folder, { recursive: true });
const deck = createPresentation({ textStylePresets: { base: { frame: { textDirection: 'horizontal' } } } });
deck.defineSlideLayout({ id: 'vertical-layout', name: 'Vertical layout', placeholders: [{
  key: 'body', kind: 'body', box: { x: 60, y: 110, width: 300, height: 240 },
  textStyle: { frame: { textDirection: 'eastAsianVertical' } },
}] });
for (const local of [false, true]) {
  const slide = deck.addSlide({ layoutId: 'vertical-layout' });
  slide.addElement({ type: 'text', box: { x: 60, y: 30, width: 500, height: 40 }, content: local ? 'LOCAL HORIZONTAL' : 'INHERITED VERTICAL' });
  slide.addElement({ type: 'text', id: local ? 'local' : 'inherited', placeholderKey: 'body', textStylePreset: 'base',
    textStylePresetReset: ['frame.textDirection'], ...(local ? { frame: { textDirection: 'horizontal' } } : {}),
    content: [{ runs: [{ text: '中文 ABC 123', style: { fontFamily: 'Arial', eastAsianFontFamily: 'Hiragino Sans GB', fontSize: 24 } }] }],
  });
}
for (const mode of ['inherit', 'materialize']) {
  await writeFile(folder + '/vertical-placeholder-' + mode + '.pptx',
    (await generatePptx(deck, { nativePlaceholderText: mode })).bytes, { flag: 'wx' });
}
