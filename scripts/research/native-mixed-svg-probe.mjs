import { mkdir, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { createPresentation } from '../../packages/core/dist/index.js';
import { renderPresentationToSvg } from '../../packages/svg-renderer/dist/index.js';

const out = resolve(process.argv[2] ?? '.tmp/native-mixed-svg-probe');
await mkdir(out, { recursive: true });
const p = createPresentation({ size: { width: 320, height: 100 }, theme: { fonts: { body: 'Arial', heading: 'Arial' } } });
p.addSlide({ id: 's', elements: [{ type: 'text', id: 'mixed', box: { x: 10, y: 10, width: 300, height: 60 }, frame: { wrap: false, margin: 0 }, content: [{ runs: [
  { text: 'Alpha ', style: { fontSize: 24, color: 'FF0000' } },
  { text: 'Beta', style: { fontSize: 24, bold: true, color: '0000FF' } },
] }] }] });
for (const [name, mixedTextMode] of [['browser', 'browser'], ['native', 'native-no-wrap']]) {
  const result = await renderPresentationToSvg(p, { mixedTextMode, portabilityDiagnostics: 'warn' });
  await writeFile(`${out}/${name}.svg`, result.slides[0].svg);
  await writeFile(`${out}/${name}-warnings.json`, `${JSON.stringify(result.warnings, null, 2)}\n`);
}
console.log(out);
