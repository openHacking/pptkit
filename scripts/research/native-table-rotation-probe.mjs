import { mkdir, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { createPresentation } from '../../packages/core/dist/index.js';
import { renderPresentationToSvg } from '../../packages/svg-renderer/dist/index.js';

const out = resolve(process.argv[2] ?? '.tmp/native-table-rotation-probe');
await mkdir(out, { recursive: true });
const p = createPresentation({ size: { width: 240, height: 200 }, theme: { fonts: { body: 'Arial', heading: 'Arial' } } });
p.addSlide({ id: 's', elements: [{ type: 'table', id: 'table', box: { x: 20, y: 20, width: 200, height: 160 }, columns: [100, 100], rows: [{ height: 160, cells: [
  { content: [{ runs: [{ text: 'Clockwise', style: { fontSize: 16, color: 'FF0000' } }] }], style: { textDirection: 'rotate90', margin: 8 } },
  { content: [{ runs: [{ text: 'Counter', style: { fontSize: 16, color: '0000FF' } }] }], style: { textDirection: 'rotate270', margin: 8 } },
] }] }] });
for (const [name, tableMode] of [['browser', 'browser'], ['native', 'native']]) {
  const result = await renderPresentationToSvg(p, { tableMode, portabilityDiagnostics: 'warn' });
  await writeFile(`${out}/${name}.svg`, result.slides[0].svg);
  await writeFile(`${out}/${name}-warnings.json`, `${JSON.stringify(result.warnings, null, 2)}\n`);
}
console.log(out);
