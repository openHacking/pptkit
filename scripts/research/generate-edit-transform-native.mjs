import { mkdir, writeFile } from 'node:fs/promises';
import { createPresentation } from '../../packages/core/dist/index.js';
import { generatePptx } from '../../packages/pptx-exporter/dist/index.js';

// A named, asymmetric shape makes rotation and reflection observable separately.
const doc = createPresentation({ id: 'transform-observation', size: { width: 720, height: 405 } });
doc.addSlide({ id: 'slide', elements: [
  { id: 'shape', name: 'Transform target', type: 'shape', shape: 'parallelogram', box: { x: 150, y: 110, width: 260, height: 120 }, style: { fill: { type: 'solid', color: '3377CC' } }, text: { content: 'Transform sample' } },
  { id: 'reference', type: 'text', box: { x: 30, y: 25, width: 500, height: 40 }, content: 'Rotation 30 degrees; horizontal flip' },
] });
doc.applyEditOperation({ schemaVersion: 'pptkit-edit-v1', documentId: doc.id, changes: [
  { type: 'element-transform', slideId: 'slide', elementId: 'shape', expected: null, value: { rotation: 30, flipH: true, flipV: null } },
] });
await mkdir('.tmp/edit-transform-native', { recursive: true });
await writeFile('.tmp/edit-transform-native/generated.pptx', (await generatePptx(doc)).bytes);
