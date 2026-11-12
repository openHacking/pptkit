import { createHash } from 'node:crypto';
import { mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { chromium } from '@playwright/test';
import { createPresentation } from '../../packages/core/dist/index.js';
import { resolveLayout } from '../../packages/layout/dist/index.js';
import { renderLayoutToSvg } from '../../packages/svg-renderer/dist/index.js';
import { renderLayoutToPdf } from '../../packages/pdf-renderer/dist/index.js';

const parent = resolve('.tmp/pdf-font-runs'); await mkdir(parent, { recursive: true });
const out = process.argv[2] === undefined ? await mkdtemp(`${parent}/run-`) : resolve(process.argv[2]); await mkdir(out, { recursive: true });
const hash = value => createHash('sha256').update(value).digest('hex');
const fontPath = 'packages/svg-renderer/test/fixtures/roboto-mono-latin-400-normal.woff2';
const fontBytes = await readFile(fontPath), family = 'PPTKit PDF Fixture Face 91237';
const face = { family, weight: 400, italic: false, format: 'woff2', base64: fontBytes.toString('base64') };
const deck = createPresentation({ size: { width: 400, height: 200 }, theme: { fonts: { body: family, heading: family } } });
for (const [id, text, linked] of [['native', 'iiii WWWW 12345', false], ['xhtml', 'Linked 12345', true]]) {
  deck.addSlide({ id, elements: [{ type: 'text', box: { x: 10, y: 10, width: 380, height: 70 }, frame: { margin: 0, wrap: false }, content: [{ runs: [{ text, style: { fontFamily: family, fontSize: 36 }, ...(linked ? { action: { type: 'url', url: 'https://example.com/font' } } : {}) }] }] }] });
}
const layout = resolveLayout(deck), svg = await renderLayoutToSvg(layout, { embeddedFonts: [face] });
const browser = await chromium.launch();
try {
  const results = [];
  for (const [name, options] of [['plain', {}], ['embedded', { fontPolicy: 'require-embedded', embeddedFonts: [face] }]]) {
    const result = await renderLayoutToPdf(layout, { browser, ...options });
    await writeFile(`${out}/${name}.pdf`, result.bytes);
    results.push({ name, sha256: hash(result.bytes), pages: result.pages, warnings: result.warnings });
  }
  let corrupt;
  try { await renderLayoutToPdf(layout, { browser, fontPolicy: 'require-embedded', embeddedFonts: [{ ...face, base64: Buffer.from('wOF2invalid-font-data').toString('base64') }] }); corrupt = 'accepted'; }
  catch (error) { corrupt = error.code ?? error.name; }
  const coverageDeck = createPresentation({ size: { width: 400, height: 200 }, theme: { fonts: { body: family, heading: family } } });
  coverageDeck.addSlide({ elements: [{ type: 'text', box: { x: 10, y: 10, width: 380, height: 70 }, content: [{ runs: [{ text: 'ABC 汉字', style: { fontFamily: family, fontSize: 36 } }] }] }] });
  const coverage = await renderLayoutToPdf(resolveLayout(coverageDeck), { browser, fontPolicy: 'require-embedded', embeddedFonts: [face] });
  await writeFile(`${out}/coverage.pdf`, coverage.bytes);
  const manifest = { producerSha256: hash(await readFile(new URL(import.meta.url))), adapterSourceSha256: hash(await readFile('packages/pdf-renderer/src/rendering/print.ts')), fontSha256: hash(fontBytes), browser: browser.version(), foreignObject: svg.slides.map(slide => slide.svg.includes('foreignObject')), results, corrupt, coverage: { sha256: hash(coverage.bytes), warnings: coverage.warnings } };
  await writeFile(`${out}/evidence.json`, `${JSON.stringify(manifest, null, 2)}\n`);
  console.log(JSON.stringify({ directory: out, ...manifest }, null, 2));
} finally { await browser.close(); }
