import { mkdir, mkdtemp, writeFile } from 'node:fs/promises';
import { createPresentation } from '../../packages/core/dist/index.js';
import { generatePptx } from '../../packages/pptx-exporter/dist/index.js';

await mkdir('.tmp', { recursive: true });
const directory = await mkdtemp('.tmp/edit-operation-native-');
const doc = createPresentation({ id: 'edit-native' });
doc.addSlide({ id: 'page', elements: [
  { id: 'headline', name: 'headline', type: 'text', box: { x: 40, y: 35, width: 320, height: 50 }, content: 'Original headline' },
  { id: 'group', name: 'group', type: 'group', box: { x: 100, y: 130, width: 200, height: 90 }, coordinateSize: { width: 200, height: 90 }, children: [
    { id: 'shape', name: 'shape', type: 'shape', shape: 'rect', box: { x: 0, y: 0, width: 200, height: 90 }, text: { content: [{ runs: [{ text: 'Original shape' }] }] } },
  ] },
] });
const result = doc.applyEditOperation({ schemaVersion: 'pptkit-edit-v1', documentId: 'edit-native', changes: [
  { type: 'text-string', slideId: 'page', elementId: 'headline', expected: 'Original headline', value: 'Edited headline' },
  { type: 'text-run', slideId: 'page', elementId: 'shape', paragraphIndex: 0, runIndex: 0, expected: 'Original shape', value: 'Edited shape' },
  { type: 'element-box', slideId: 'page', elementId: 'group', expected: { x: 100, y: 130, width: 200, height: 90 }, value: { x: 180, y: 150, width: 200, height: 90 } },
] });
await writeFile(`${directory}/source.pptx`, (await generatePptx(doc)).bytes);
await writeFile(`${directory}/inverse.json`, JSON.stringify(result.inverse, null, 2) + '\n');
console.log(`${directory}/source.pptx`);
