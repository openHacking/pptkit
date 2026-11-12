import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { createPresentation } from '../../packages/core/dist/index.js';
import { renderPresentationToSvg } from '../../packages/svg-renderer/dist/index.js';

const out = resolve(process.argv[2] ?? '.tmp/embedded-font-mupdf-probe');
await mkdir(out, { recursive: true });
const family = 'PPTKit Embedded MuPDF Alias';
const d = createPresentation({ size: { width: 340, height: 80 }, theme: { fonts: { body: family, heading: family } } });
d.addSlide({ id: 's', elements: [{ type: 'text', id: 't', box: { x: 4, y: 4, width: 330, height: 60 }, frame: { margin: 0, wrap: false }, content: [{ runs: [{ text: 'iiii WWWW 12345', style: { fontFamily: family, fontSize: 36 } }] }] }] });
const font = { family, weight: 400, italic: false, format: 'woff2', base64: (await readFile('packages/svg-renderer/test/fixtures/roboto-mono-latin-400-normal.woff2')).toString('base64') };
for (const [name, options] of [['plain', {}], ['embedded', { embeddedFonts: [font] }]]) {
  const result = await renderPresentationToSvg(d, options);
  await writeFile(`${out}/${name}.svg`, result.slides[0].svg);
  await writeFile(`${out}/${name}-warnings.json`, `${JSON.stringify(result.warnings, null, 2)}\n`);
}
console.log(out);
