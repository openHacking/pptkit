import { chromium, firefox, webkit } from '@playwright/test';
import assert from 'node:assert/strict';
import { createPresentation } from '../../packages/core/dist/index.js';
import { renderPresentationToSvg } from '../../packages/svg-renderer/dist/index.js';

const sources = [
  'ABC (שלום 123) XYZ',
  'ABC (שלום 123) XYZ\r\n  مرحبا 123 XYZ\nEnd',
  '\r\n\r  \nABC שלום 123\n\n',
  'ABC שלום 123\r\n\nمرحبا 456 XYZ',
  'ABC (words 123) XYZ',
];
const deck = createPresentation();
const slide = deck.addSlide();
for (const [i, text] of sources.entries()) slide.addElement({ type: 'text', id: `t${i}`, box: { x: 0, y: i * 220, width: 190, height: 200 }, frame: { margin: 0, wrap: true }, content: [{ ...(i === 3 ? { style: { bullet: { type: 'bullet', character: '•' } } } : {}), runs: [{ text, style: { fontFamily: 'Arial', fontSize: 32 } }] }] });
const svg = (await renderPresentationToSvg(deck)).slides[0].svg;
for (const [name, browserType] of [['chromium', chromium], ['firefox', firefox], ['webkit', webkit]]) {
  const browser = await browserType.launch({ headless: true });
  try {
    const page = await browser.newPage();
    await page.setContent(svg);
    const samples = await page.evaluate(ids => ids.map(id => {
      const node = document.querySelector(`[data-pptkit-element-id="${id}"]`);
      const range = document.createRange(); range.selectNodeContents(node);
      const selection = getSelection(); selection.removeAllRanges(); selection.addRange(range);
      const selected = selection.toString(); selection.removeAllRanges();
      const lines = [...node.querySelectorAll('[data-pptkit-text-line]')];
      const positions = target => [...target.querySelectorAll('tspan[font-family]')].flatMap(span =>
        Array.from({ length: span.getNumberOfChars() }, (_, i) => {
          const box = span.getExtentOfChar(i); return [[box.x, box.y, box.width, box.height]];
        }));
      const matches = lines.map(line => {
        const original = positions(line), clone = line.cloneNode(true);
        for (const span of clone.querySelectorAll('tspan:not([font-family])')) span.remove();
        node.appendChild(clone); const control = positions(clone); clone.remove();
        return JSON.stringify(original) === JSON.stringify(control);
      });
      const extraEmptyLine = [...node.querySelectorAll('span[aria-hidden="true"]')].some(span => span.textContent === '\u200b');
      return { dom: node.textContent, selected, extraEmptyLine, mode: node.tagName.toLowerCase(), lines: lines.length, matches };
    }), sources.map((_, index) => `t${index}`));
    console.log(JSON.stringify({ browser: name, version: browser.version(), samples }));
    for (const [i, sample] of samples.entries()) {
      assert.equal(sample.dom, `${i === 3 ? '• ' : ''}${sources[i]}${sample.extraEmptyLine ? '\u200b' : ''}`, `${name}: source text ${i}`);
      if (sample.mode === 'g') assert.ok(sample.matches.length > 1 && sample.matches.every(Boolean), `${name}: glyph positions ${i}`);
      else assert.equal(sample.mode, 'foreignobject', `${name}: paragraph fallback ${i}`);
    }
  } finally { await browser.close(); }
}
