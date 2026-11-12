import { createHash } from 'node:crypto';
import { mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { importPptx } from '../packages/pptx-parser/dist/index.js';

const original = await readFile('.tmp/native-scatter-research/powerpoint-grywvvg1/scatter-research-round-1.pptx');
const session = await importPptx(original, { projection: 'content', chartEditSaveProfile: 'existing-values-v2' });
let saved;
try {
  if (session.report.chartValues.length !== 60 || session.report.chartValues.some(b => b.status !== 'bound')) throw new Error('Unexpected source bindings.');
  for (const slide of session.document.slides) {
    const chart = structuredClone(slide.elements.find(e => e.type === 'chart'));
    if (chart.series[0].yValues[1] !== null) throw new Error('Expected blank source point.');
    chart.series[0].xValues[0] = -6; chart.series[0].yValues[2] = -3;
    slide.replaceElement(chart.id, chart);
  }
  const plan = await session.prepareSave({ mode: 'edit' });
  if (plan.status !== 'ready') throw new Error(JSON.stringify(plan.report));
  saved = await session.save(plan.plan);
} finally { session.dispose(); }
await mkdir('.tmp/native-scatter-sparse-save', { recursive: true });
const dir = path.resolve(await mkdtemp('.tmp/native-scatter-sparse-save/powerpoint-'));
const sha = bytes => createHash('sha256').update(bytes).digest('hex');
await writeFile(path.join(dir, 'source.pptx'), original, { flag: 'wx' });
await writeFile(path.join(dir, 'edited.pptx'), saved.bytes, { flag: 'wx' });
await writeFile(path.join(dir, 'manifest.json'), JSON.stringify({ sourceSha256: sha(original), editedSha256: sha(saved.bytes), changes: saved.report.changes.map(change => change.chartValue), expected: { RED: [[-6, 4], [-3, null], [2, -3], [8, 6]], BLUE: [[-7, -2], [-1, 4], [3, 2], [7, -6]] } }, null, 2) + '\n', { flag: 'wx' });
console.log(dir);
