import { mkdir, writeFile } from 'node:fs/promises';
import { createPresentation } from '../packages/core/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';
import { png } from '../packages/pptx-parser/test/image-fixture.mjs';

const source = createPresentation(), destination = createPresentation();
source.registerAsset({ id: 'source-picture', kind: 'image', mimeType: 'image/png', source: { type: 'bytes', value: png }, width: 8, height: 4 });
source.addSlide({ id: 'source', notes: 'Copied page notes', elements: [
  { id: 'title', type: 'text', content: 'Copied page', box: { x: 80, y: 45, width: 400, height: 65 }, action: { type: 'slide', slideId: 'source' } },
  { id: 'picture', type: 'image', assetId: 'source-picture', box: { x: 100, y: 150, width: 280, height: 140 } },
  { id: 'shape', type: 'shape', shape: 'rect', box: { x: 530, y: 170, width: 160, height: 100 } },
  { id: 'connector', type: 'connector', start: { elementId: 'picture', anchor: 'right' }, end: { elementId: 'shape', anchor: 'left' } },
] });
destination.addSlide({ id: 'control', notes: 'Control page notes', elements: [{ id: 'control-text', type: 'text', content: 'Control page', box: { x: 80, y: 45, width: 400, height: 65 } }] });
destination.copySlideFrom(source, 'source', { layoutId: null, theme: 'destination', geometry: 'preserve', customData: 'preserve', assets: { deduplicate: 'content', external: 'reject' } });
source.removeSlide('source');
const output = '.tmp/cross-document-copy-native';
await mkdir(output, { recursive: true });
await writeFile(`${output}/input.pptx`, (await generatePptx(destination)).bytes);
