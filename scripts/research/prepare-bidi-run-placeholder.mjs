import { mkdir, writeFile } from 'node:fs/promises';
import { createPresentation } from '../../packages/core/dist/index.js';
import { generatePptx } from '../../packages/pptx-exporter/dist/index.js';

const folder = process.argv[2] ?? '.tmp/bidi-run-placeholder-marked';
await mkdir(folder, { recursive: true });
const deck = createPresentation({ textStylePresets: { reset: { run: { rightToLeft: false } } } });
const box = { x: 60, y: 80, width: 680, height: 130 };
for (const value of [true, false]) {
  const layoutId = value ? 'true-layout' : 'false-layout';
  deck.defineSlideLayout({ id: layoutId, name: layoutId, placeholders: [
    { key: 'body', kind: 'body', box, textStyle: { run: { rightToLeft: value } } },
  ] });
  const slide = deck.addSlide({ layoutId });
  slide.addElement({ type: 'text', id: `page-marker-${value}`, box: { x: 60, y: 300, width: 500, height: 60 },
    content: value ? 'PAGE ONE' : 'PAGE TWO' });
  slide.addElement({ type: 'text', id: layoutId, placeholderKey: 'body', textStylePreset: 'reset', textStylePresetReset: ['rightToLeft'],
    content: [{ runs: [
      { text: 'ABC ', style: { rightToLeft: false } }, { text: 'مرحبا - 123' }, { text: ' XYZ', style: { rightToLeft: false } },
      { text: ' שלום - 456', style: { rightToLeft: !value } },
    ] }], frame: { wrap: false },
  });
}
for (const mode of ['inherit', 'materialize']) {
  const { bytes } = await generatePptx(deck, { nativePlaceholderText: mode });
  await writeFile(`${folder}/rtl-${mode}.pptx`, bytes);
}
