import assert from 'node:assert/strict';
import { readFile, writeFile, mkdir } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { execFileSync } from 'node:child_process';
import { chromium } from '@playwright/test';
import { createPresentation } from '../../packages/core/dist/index.js';
import { generatePptx } from '../../packages/pptx-exporter/dist/index.js';

const path = process.argv[2] ?? '/System/Library/Fonts/Supplemental/NotoSansMongolian-Regular.ttf';
const folder = process.argv[3] ?? '.tmp/vertical-writing/mongolian';
await mkdir(folder, { recursive: true });
const variants = process.argv[4] === '--variants';
const bytes = await readFile(path), text = variants ? 'ᠠ\u180B ᠠ\u180C ᠠ\u180D ᠠ\u180Eᠠ' : 'ᠮᠣᠩᠭᠣᠯ';
const sha256 = createHash('sha256').update(bytes).digest('hex');
const shape = direction => JSON.parse(execFileSync('hb-shape', ['--output-format=json', '--font-size=36000', '--direction=' + direction, '--script=Mong', '--language=mn', path, text], { encoding: 'utf8' }));
const hb = { ltr: shape('ltr'), ttb: shape('ttb') };
for (const sequence of Object.values(hb)) {
  assert.ok(sequence.every(g => g.g !== '.notdef'));
  if (!variants) {
    assert.match(sequence[0].g, /init/);
    assert.match(sequence.at(-1).g, /fina/);
  }
}
const alias = 'PPTKit Mongolian ' + sha256.slice(0, 12);
const native = createPresentation();
const slide = native.addSlide();
slide.addElement({ type: 'text', box: { x: 60, y: 20, width: 400, height: 40 }, content: 'MONGOLIAN VERTICAL' });
slide.addElement({ type: 'text', box: { x: 60, y: 100, width: 180, height: 300 }, frame: { textDirection: 'mongolianVertical', wrap: false },
  content: [{ runs: [{ text, style: { fontFamily: 'Arial', complexScriptFontFamily: 'Noto Sans Mongolian', fontSize: 36, language: 'mn-Mong' } }] }] });
await writeFile(folder + '/mongolian-native.pptx', (await generatePptx(native)).bytes);

const imports = Object.fromEntries(await Promise.all(['core', 'layout', 'svg-renderer'].map(async name => [
  '@pptkit/' + name, 'data:text/javascript;base64,' + (await readFile(`packages/${name}/dist/browser/index.js`)).toString('base64'),
])));
const browser = await chromium.launch();
try {
  const page = await browser.newPage({ viewport: { width: 400, height: 450 } });
  await page.route('**/*', route => route.abort());
  await page.setContent(`<script type="importmap">${JSON.stringify({ imports })}</script><style>body{margin:0;background:white}</style>`);
  const advance = await page.evaluate(async ({ alias, font, text }) => {
    const { createPresentation } = await import('@pptkit/core');
    const { renderPresentationToSvg } = await import('@pptkit/svg-renderer');
    const face = await new FontFace(alias, 'url(data:font/ttf;base64,' + font + ')').load();
    document.fonts.add(face);
    const deck = createPresentation({ size: { width: 400, height: 450 } });
    deck.addSlide().addElement({ type: 'text', box: { x: 60, y: 60, width: 200, height: 350 },
      frame: { textDirection: 'mongolianVertical', wrap: false },
      content: [{ runs: [{ text, style: { complexScriptFontFamily: alias, fontSize: 36, language: 'mn-Mong' } }] }] });
    document.body.insertAdjacentHTML('beforeend', (await renderPresentationToSvg(deck)).slides[0].svg);
    const span = document.querySelector('foreignObject span[lang]');
    if (!span || span.textContent !== text) throw new Error('Source run lost or split');
    span.id = 'probe';
    const range = document.createRange(); range.selectNodeContents(span);
    return range.getBoundingClientRect().height;
  }, { alias, font: bytes.toString('base64'), text });
  const cdp = await page.context().newCDPSession(page);
  await cdp.send('DOM.enable'); await cdp.send('CSS.enable');
  const { root } = await cdp.send('DOM.getDocument');
  const { nodeId } = await cdp.send('DOM.querySelector', { nodeId: root.nodeId, selector: '#probe' });
  const { fonts } = await cdp.send('CSS.getPlatformFontsForNode', { nodeId });
  assert.equal(fonts.length, 1);
  assert.equal(fonts[0].isCustomFont, true);
  assert.equal(fonts[0].postScriptName, 'NotoSansMongolian-Regular');
  if (!variants) assert.equal(fonts[0].glyphCount, hb.ltr.length);
  else assert.ok(fonts[0].glyphCount >= hb.ltr.filter(g => g.ax !== 0).length);
  const expectedAdvance = hb.ltr.reduce((sum, glyph) => sum + glyph.ax, 0) / 1000;
  assert.ok(Math.abs(advance - expectedAdvance) < 0.05, { advance, expectedAdvance });
  await page.screenshot({ path: folder + '/mongolian-browser.png' });
  const result = { text, sha256, chromium: browser.version(), fonts, advance, expectedAdvance, hb };
  await writeFile(folder + '/mongolian-browser.json', JSON.stringify(result, null, 2) + '\n');
  console.log(JSON.stringify(result));
} finally { await browser.close(); }
