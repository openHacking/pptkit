import { mkdir, writeFile, copyFile } from 'node:fs/promises';
import { createPresentation } from '../packages/core/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';
import { importPptx } from '../packages/pptx-parser/dist/index.js';
import { createChartSaveFixture } from '../tests/fixtures/chart-save.mjs';

const directory = process.argv[2];
if (!directory) throw new Error('Supply a new output directory.');
await mkdir(directory, { recursive: false });
const deck = createChartSaveFixture(createPresentation), slide = deck.slides[0];
slide.addElement({ id: 'number-field', type: 'text', box: { x: 40, y: 470, width: 600, height: 40 }, content: [{ runs: [
  { text: 'Page: ' }, { text: 'STALE', field: { id: '11111111-2222-3333-4444-555555555555', kind: 'slide-number' } }, { text: ' / note' },
] }] });
const source = (await generatePptx(deck)).bytes;
await writeFile(`${directory}/source.pptx`, source, { flag: 'wx' });
const session = await importPptx(source, { projection: 'content', chartEditSaveProfile: 'existing-values-v1', firstSlideNumberSaveProfile: 'attribute-v2', editSaveProfile: 'text-runs-v3', fieldTextProfile: 'slide-number-v1' });
try {
  const imported = session.document.slides[0], chart = structuredClone(imported.elements[1]), text = structuredClone(imported.elements[2]);
  if (text.content[0].runs[1].field?.kind !== 'slide-number') throw new Error('Field not projected');
  text.content[0].runs[0].text = 'Edited page: '; imported.replaceElement(text.id, text);
  chart.series[0].values[1] = 7; imported.replaceElement(chart.id, chart); session.document.setFirstSlideNumber(17);
  const prepared = await session.prepareSave({ mode: 'edit' });
  if (prepared.status !== 'ready' || prepared.plan.report.changes.length !== 3) throw new Error(`Unexpected plan: ${JSON.stringify(prepared.report)}`);
  await writeFile(`${directory}/candidate.pptx`, (await session.save(prepared.plan)).bytes, { flag: 'wx' });
  await copyFile(`${directory}/candidate.pptx`, `${directory}/native-copy.pptx`);
} finally { session.dispose(); }
