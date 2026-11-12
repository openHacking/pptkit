import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { importPptx } from '../../packages/pptx-parser/dist/index.js';

const directory = process.argv[2] ?? '.tmp/series-name-native-reproduction';
await mkdir(directory, { recursive: true });
let bytes = await readFile('tests/fixtures/chart-label-source.pptx');
for (const [index, value] of ['Native revised', 'Native second', 'Native third'].entries()) {
  const session = await importPptx(bytes, { projection: 'content', chartEditSaveProfile: 'existing-values-and-series-names-v5' });
  try {
    const slide = session.document.slides[0], chart = structuredClone(slide.elements[0]);
    chart.series[0].name = value; slide.replaceElement(chart.id, chart);
    const plan = await session.prepareSave({ mode: 'edit' });
    if (plan.status !== 'ready') throw new Error(JSON.stringify(plan.report));
    bytes = (await session.save(plan.plan)).bytes;
    await writeFile(`${directory}/round-${index + 1}.pptx`, bytes);
  } finally { session.dispose(); }
}
