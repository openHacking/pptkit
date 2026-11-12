import { mkdir, writeFile } from 'node:fs/promises';
import { createPresentation } from '../../packages/core/dist/index.js';
import { generatePptx } from '../../packages/pptx-exporter/dist/index.js';

const doc = createPresentation({ id: 'action-observation', size: { width: 720, height: 405 } });
doc.addSlide({ id: 'first', elements: [
  { id: 'title1', type: 'text', box: { x: 40, y: 30, width: 500, height: 45 }, content: 'FIRST: click blue shape for THIRD' },
  { id: 'button', name: 'Go to third', type: 'shape', shape: 'rect', box: { x: 180, y: 140, width: 300, height: 100 }, style: { fill: { type: 'solid', color: '3377CC' } }, text: { content: 'Go to THIRD' } },
] });
doc.addSlide({ id: 'second', elements: [{ id: 'title2', type: 'text', box: { x: 40, y: 30, width: 500, height: 45 }, content: 'SECOND: table action destination' }] });
doc.addSlide({ id: 'third', elements: [
  { id: 'title3', type: 'text', box: { x: 40, y: 30, width: 600, height: 45 }, content: 'THIRD: choose a link' },
  { id: 'run', type: 'text', box: { x: 100, y: 120, width: 500, height: 55 }, content: [{ runs: [{ text: 'Text link to FIRST', style: { fontSize: 26 } }] }] },
  { id: 'table', type: 'table', box: { x: 100, y: 220, width: 500, height: 70 }, columns: [500], rows: [{ height: 70, cells: [{ content: [{ runs: [{ text: 'Table link to SECOND', style: { fontSize: 26 } }] }] }] }] },
] });
doc.applyEditOperation({ schemaVersion: 'pptkit-edit-v1', documentId: doc.id, changes: [
  { type: 'element-action', slideId: 'first', elementId: 'button', expected: null, value: { type: 'slide', slideId: 'third', tooltip: 'Third slide' } },
  { type: 'text-run-action', slideId: 'third', elementId: 'run', paragraphIndex: 0, runIndex: 0, expected: null, value: { type: 'slide', slideId: 'first', tooltip: 'First slide' } },
  { type: 'table-cell-text-run-action', slideId: 'third', elementId: 'table', rowIndex: 0, cellIndex: 0, paragraphIndex: 0, runIndex: 0, expected: null, value: { type: 'slide', slideId: 'second', tooltip: 'Second slide' } },
] });
await mkdir('.tmp/edit-action-native', { recursive: true });
await writeFile('.tmp/edit-action-native/generated.pptx', (await generatePptx(doc)).bytes);
