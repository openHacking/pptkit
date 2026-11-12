import { createHash } from 'node:crypto';
import { mkdir, mkdtemp, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { createPresentation } from '../packages/core/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';

await mkdir('.tmp/native-connector-interaction', { recursive: true });
if (process.argv.length > 3 || (process.argv[2] !== undefined && process.argv[2] !== '--elbow')) throw new Error('Usage: node scripts/prepare-native-connector-interaction.mjs [--elbow]');
const elbow = process.argv[2] === '--elbow';
const directory = path.resolve(await mkdtemp('.tmp/native-connector-interaction/run-'));
const deck = createPresentation({ id: 'connector-interaction', size: { width: 960, height: 540 } });
const slide = deck.addSlide({ id: 'slide' });
slide.addElement({ type: 'connector', id: 'link', name: 'link', start: { elementId: 'left', anchor: 'right' }, end: { elementId: 'right', anchor: 'left' }, ...(elbow ? { route: [{ x: 600, y: 260 }] } : {}), style: { paint: { type: 'solid', color: 'C32626' }, width: 3, endArrow: 'triangle' } });
for (const [id, x, color] of [['left', 150, '4472C4'], ['right', 600, '70AD47']]) slide.addElement({ type: 'shape', id, name: id, shape: 'rect', box: { x, y: elbow && id === 'right' ? 340 : 210, width: 150, height: 100 }, style: { fill: { type: 'solid', color } } });
const result = await generatePptx(deck, { connectorMode: 'native' });
if (result.warnings.length) throw new Error(JSON.stringify(result.warnings));
await writeFile(path.join(directory, 'source.pptx'), result.bytes, { flag: 'wx' });
await writeFile(path.join(directory, 'manifest.json'), JSON.stringify({ sha256: createHash('sha256').update(result.bytes).digest('hex'), scope: `Two rectangles joined by one native ${elbow ? 'single-elbow' : 'straight'} connector; no native interaction has yet been verified.` }, null, 2) + '\n', { flag: 'wx' });
console.log(directory);
