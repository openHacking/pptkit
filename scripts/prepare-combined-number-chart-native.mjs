import { mkdir, writeFile, copyFile } from 'node:fs/promises';
import { createPresentation } from '../packages/core/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';
import { importPptx } from '../packages/pptx-parser/dist/index.js';
import { createChartSaveFixture } from '../tests/fixtures/chart-save.mjs';

const directory = process.argv[2];
if (!directory) throw new Error('Supply a new output directory.');
await mkdir(directory, { recursive: false });
const source = (await generatePptx(createChartSaveFixture(createPresentation))).bytes;
await writeFile(`${directory}/source.pptx`, source, { flag: 'wx' });
const session = await importPptx(source, { projection: 'content', chartEditSaveProfile: 'existing-values-v1', firstSlideNumberSaveProfile: 'attribute-v2', editSaveProfile: 'text-runs-v3' });
try {
  const slide = session.document.slides[0], title = structuredClone(slide.elements[0]), chart = structuredClone(slide.elements[1]);
  title.content[0].runs[0].text = 'Combined chart and page number'; slide.replaceElement(title.id, title);
  chart.series[0].values[1] = 7; slide.replaceElement(chart.id, chart); session.document.setFirstSlideNumber(17);
  const prepared = await session.prepareSave({ mode: 'edit' });
  if (prepared.status !== 'ready' || prepared.plan.report.changes.length !== 3) throw new Error(`Unexpected save plan: ${JSON.stringify(prepared.report)}`);
  await writeFile(`${directory}/candidate.pptx`, (await session.save(prepared.plan)).bytes, { flag: 'wx' });
  await copyFile(`${directory}/candidate.pptx`, `${directory}/native-copy.pptx`);
} finally { session.dispose(); }
