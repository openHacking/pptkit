// Named first-open probe; control.png is made from the 3x2 RGB PPM input.
import { readFile, writeFile } from 'node:fs/promises';
import { createPresentation } from '../packages/core/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';

const directory = process.argv[2] ?? '.tmp/native-jpeg-12';
const presentation = createPresentation(), slide = presentation.addSlide();
const items = [
  ['12-bit sequential', 'tests/fixtures/jpeg-12bit-sequential.jpg'],
  ['12-bit progressive', 'tests/fixtures/jpeg-12bit-progressive.jpg'],
  ['PNG control', `${directory}/control.png`],
];
for (const [index, [label, file]] of items.entries()) {
  const id = `a${index}`;
  presentation.registerAsset({ id, kind: 'image', source: { type: 'bytes', value: new Uint8Array(await readFile(file)) } });
  slide.addElement({ type: 'text', content: label, box: { x: 30 + index * 300, y: 60, width: 280, height: 50 } });
  slide.addElement({ type: 'image', assetId: id, fit: 'stretch', box: { x: 30 + index * 300, y: 130, width: 260, height: 260 } });
}
const output = await generatePptx(presentation);
await writeFile(`${directory}/probe.pptx`, output.bytes);
