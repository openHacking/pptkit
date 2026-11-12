import { mkdir, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { createPresentation } from '../../packages/core/dist/index.js';
import { renderPresentationToSvg } from '../../packages/svg-renderer/dist/index.js';

const out = resolve(process.argv[2] ?? '.tmp/native-table-svg-probe');
await mkdir(out, { recursive: true });
const p = createPresentation({ size: { width: 240, height: 120 }, theme: { fonts: { body: 'Arial', heading: 'Arial' } } });
const run = text => [{ runs: [{ text, style: { fontSize: 16 } }] }];
p.addSlide({ id: 's', elements: [{ type: 'table', id: 'table', box: { x: 20, y: 20, width: 200, height: 80 }, columns: [100, 100], rows: [
  { height: 40, cells: [{ colSpan: 2, content: run('Header'), style: { margin: 4, fill: { type: 'solid', color: 'FFFF00' } } }] },
  { height: 40, cells: [{ content: run('Left'), style: { margin: 4, fill: { type: 'solid', color: '00FFFF' } } }, { content: run('Right'), style: { margin: 4, fill: { type: 'solid', color: 'FFB0B0' } } }] },
] }] });
for (const [name, tableMode] of [['browser', 'browser'], ['native', 'native']]) {
  const result = await renderPresentationToSvg(p, { tableMode, portabilityDiagnostics: 'warn' });
  await writeFile(`${out}/${name}.svg`, result.slides[0].svg);
  await writeFile(`${out}/${name}-warnings.json`, `${JSON.stringify(result.warnings, null, 2)}\n`);
}
console.log(out);
