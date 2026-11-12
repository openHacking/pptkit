import { createHash } from 'node:crypto';
import { mkdir, mkdtemp, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { createPresentation } from '../packages/core/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';

await mkdir('.tmp/native-connectors', { recursive: true });
const directory = path.resolve(await mkdtemp('.tmp/native-connectors/powerpoint-elbow-'));
const deck = createPresentation({ id: 'native-elbow-controls', size: { width: 960, height: 540 } });
const directions = [[1, 1], [-1, 1], [1, -1], [-1, -1]];
for (const kind of ['single', 'double']) {
  const slide = deck.addSlide({ id: kind });
  slide.addElement({ id: `${kind}-title`, type: 'text', box: { x: 30, y: 20, width: 800, height: 40 }, frame: { margin: 0 }, content: `NATIVE ${kind.toUpperCase()} ELBOWS — four directions` });
  for (const [index, [sx, sy]] of directions.entries()) {
    const cx = index % 2 ? 680 : 220, cy = index < 2 ? 160 : 360;
    const start = { x: cx - sx * 100, y: cy - sy * 50 }, end = { x: cx + sx * 100, y: cy + sy * 50 };
    const middle = start.x + (end.x - start.x) * .25;
    const route = kind === 'single' ? [{ x: end.x, y: start.y }] : [{ x: middle, y: start.y }, { x: middle, y: end.y }];
    slide.addElement({ id: `${kind}-${index}`, name: `${kind}-${index}`, type: 'connector', start, end, route, style: { paint: { type: 'solid', color: 'C32626' }, width: 3, beginArrow: 'oval', endArrow: 'triangle' } });
    slide.addElement({ id: `${kind}-${index}-label`, type: 'text', box: { x: cx - 150, y: cy + 70, width: 300, height: 30 }, frame: { margin: 0 }, content: `${kind} ${index + 1} dx=${sx} dy=${sy}` });
  }
}
const result = await generatePptx(deck, { connectorMode: 'native' });
if (result.warnings.length) throw new Error(JSON.stringify(result.warnings));
await writeFile(path.join(directory, 'elbows.pptx'), result.bytes, { flag: 'wx' });
await writeFile(path.join(directory, 'manifest.json'), JSON.stringify({ sourceSha256: createHash('sha256').update(result.bytes).digest('hex'), directions, kinds: ['single', 'double'], adjustment: 25000 }, null, 2) + '\n', { flag: 'wx' });
console.log(directory);
