import { writeFile } from 'node:fs/promises';
import { createPresentation } from '../../packages/core/dist/index.js';
import { generatePptx } from '../../packages/pptx-exporter/dist/index.js';

for (const chartType of ['scatter', 'bubble']) {
  const document = createPresentation({ id: `${chartType}-name-source` });
  document.addSlide({ id: 'slide', elements: [{ id: 'chart', type: 'chart', chartType, box: { x: 0, y: 0, width: 300, height: 200 }, series: [{ name: 'Source & name', xValues: [1, 2], yValues: [3, 4], ...(chartType === 'bubble' ? { sizes: [5, 6] } : {}) }] }] });
  await writeFile(`tests/fixtures/${chartType}-name-source.pptx`, (await generatePptx(document)).bytes);
}
