import { mkdir, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { createPresentation } from '../../packages/core/dist/index.js';
import { renderPresentationToSvg } from '../../packages/svg-renderer/dist/index.js';

const out = resolve(process.argv[2] ?? '.tmp/native-wrapped-svg-probe');
await mkdir(out, { recursive: true });
const p = createPresentation({ size: { width: 180, height: 120 }, theme: { fonts: { body: 'Arial', heading: 'Arial' } } });
p.addSlide({ id: 's', elements: [{ type: 'text', id: 'text', box: { x: 10, y: 10, width: 75, height: 100 }, frame: { wrap: true, margin: 0 }, content: [{ runs: [
  { text: 'Alpha ', style: { fontSize: 20, color: 'FF0000' } },
  { text: 'Beta Gamma', style: { fontSize: 20, color: '0000FF', bold: true } },
] }] }] });
for (const [name, mixedTextMode] of [['browser', 'browser'], ['native', 'native']]) {
  const result = await renderPresentationToSvg(p, { mixedTextMode, portabilityDiagnostics: 'warn' });
  await writeFile(`${out}/${name}.svg`, result.slides[0].svg);
  await writeFile(`${out}/${name}-warnings.json`, `${JSON.stringify(result.warnings, null, 2)}\n`);
}
console.log(out);
