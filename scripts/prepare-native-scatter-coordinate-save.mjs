import { createHash } from 'node:crypto';
import { mkdir, mkdtemp, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { createPresentation } from '../packages/core/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';
import { importPptx } from '../packages/pptx-parser/dist/index.js';

const p = createPresentation({ id: 'scatter-coordinate-save', size: { width: 960, height: 540 } });
const slide = p.addSlide({ id: 'scatter' });
slide.addElement({ type: 'text', id: 'title', box: { x: 30, y: 10, width: 800, height: 45 }, content: 'Scatter coordinate preservation' });
slide.addElement({ type: 'chart', chartType: 'scatter', id: 'chart', box: { x: 40, y: 60, width: 800, height: 420 }, legend: { visible: false },
  axes: { x: { scale: { min: -10, max: 10, majorUnit: 5 } }, y: { scale: { min: -10, max: 10, majorUnit: 5 } } },
  series: [{ name: 'RED', xValues: [-8, 2, 8], yValues: [4, -4, 6], color: '#FF0000', marker: { shape: 'circle', size: 10 }, line: { paint: { type: 'solid', color: '#FF0000' }, width: 2 } },
    { name: 'BLUE', xValues: [7, -1, 3], yValues: [-6, 4, 2], color: '#0000FF', marker: { shape: 'circle', size: 10 }, line: { paint: { type: 'solid', color: '#0000FF' }, width: 2 } }] });
const source = (await generatePptx(p)).bytes;
const session = await importPptx(source, { projection: 'content', chartEditSaveProfile: 'existing-values-v1' });
let saved;
try {
  const owner = session.document.slides[0], chart = structuredClone(owner.elements.find(e => e.type === 'chart'));
  chart.series[0].xValues[0] = -3; chart.series[0].yValues[0] = 8; chart.series[1].yValues[1] = -7;
  owner.replaceElement(chart.id, chart);
  const plan = await session.prepareSave({ mode: 'edit' });
  if (plan.status !== 'ready') throw new Error(JSON.stringify(plan.report));
  saved = await session.save(plan.plan);
} finally { session.dispose(); }
await mkdir('.tmp/native-scatter-coordinate-save', { recursive: true });
const dir = path.resolve(await mkdtemp('.tmp/native-scatter-coordinate-save/powerpoint-'));
const sha = bytes => createHash('sha256').update(bytes).digest('hex');
await writeFile(path.join(dir, 'source.pptx'), source, { flag: 'wx' });
await writeFile(path.join(dir, 'edited.pptx'), saved.bytes, { flag: 'wx' });
await writeFile(path.join(dir, 'manifest.json'), JSON.stringify({ sourceSha256: sha(source), editedSha256: sha(saved.bytes), expected: { RED: [[-3, 8], [2, -4], [8, 6]], BLUE: [[7, -6], [-1, -7], [3, 2]] }, changes: saved.report.changes.map(change => change.chartValue) }, null, 2) + '\n', { flag: 'wx' });
console.log(dir);
