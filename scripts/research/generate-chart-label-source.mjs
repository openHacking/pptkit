import { mkdir, writeFile } from 'node:fs/promises';
import { createPresentation } from '../../packages/core/dist/index.js';
import { generatePptx } from '../../packages/pptx-exporter/dist/index.js';

const doc = createPresentation({ id: 'label-binding-research' });
doc.addSlide({ id: 'page', elements: [{
  id: 'chart', type: 'chart', chartType: 'bar', box: { x: 40, y: 40, width: 500, height: 300 },
  categories: [' Q1 ', 'literal _x0041_', '中文 & two'],
  series: [{ name: 'First _x005F_', values: [1, 2, 3] }, { name: 'Second & series', values: [4, 5, 6] }],
}] });
await mkdir('.tmp/chart-label-source', { recursive: true });
await writeFile('.tmp/chart-label-source/source.pptx', (await generatePptx(doc)).bytes);
