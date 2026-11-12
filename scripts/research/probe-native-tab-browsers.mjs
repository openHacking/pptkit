import assert from 'node:assert/strict';
import { chromium, firefox, webkit } from '@playwright/test';
import { createPresentation } from '../../packages/core/dist/index.js';
import { renderPresentationToSvg } from '../../packages/svg-renderer/dist/index.js';
const deck = createPresentation(), slide = deck.addSlide(), expected = new Map();
for (const kind of ['text', 'shape']) for (const align of ['left', 'center', 'right', 'decimal']) {
  const source = align === 'decimal' ? 'P\t12.34\r\nQ\t56.78\rR\t90.12\nS\t34.56' : 'P\t1234\r\nQ\t5678\rR\t9012\nS\t3456';
  const id = `${kind}-${align}`, content = [{ style: { tabStops: [{ position: 120, align }] }, runs: [{ text: source, style: { fontSize: 20 } }] }];
  const box = { x: 0, y: 0, width: 300, height: 200 }, frame = { margin: 0, wrap: false };
  slide.addElement(kind === 'text' ? { id, type: kind, box, frame, content } : { id, type: kind, shape: 'rect', box, text: { frame, content } });
  expected.set(id, source);
}
const svg = (await renderPresentationToSvg(deck)).slides[0].svg;
for (const [name, engine] of [['chromium', chromium], ['firefox', firefox], ['webkit', webkit]]) {
  const browser = await engine.launch({ headless: true });
  try {
    const page = await browser.newPage(); await page.setContent(svg);
    const actual = await page.evaluate(() => [...document.querySelectorAll('[data-pptkit-element-id]')].map(node => {
      const positions = text => [...text.querySelectorAll('tspan[font-family]')].flatMap(span => Array.from({ length: span.getNumberOfChars() }, (_, index) => { const b = span.getExtentOfChar(index); return [[b.x,b.y,b.width,b.height]]; }));
      const chunks = [...node.querySelectorAll('text[data-pptkit-tab-part]')];
      const matches = chunks.map(chunk => {
        const original = positions(chunk), clone = chunk.cloneNode(true);
        clone.querySelectorAll('tspan:not([font-family])').forEach(child => child.remove()); node.appendChild(clone);
        const control = positions(clone); clone.remove(); return JSON.stringify(original) === JSON.stringify(control);
      });
      return { id: node.getAttribute('data-pptkit-element-id'), text: node.textContent, chunks: chunks.length, matches };
    }));
    for (const item of actual) { assert.equal(item.text, expected.get(item.id), `${name}: ${item.id} text`); assert.ok(item.chunks >= 8 && item.matches.every(Boolean), `${name}: ${item.id} geometry`); }
    assert.equal(actual.length, expected.size);
    console.log(JSON.stringify({ browser: name, version: browser.version(), cases: actual.length, chunks: actual.reduce((n, item) => n + item.chunks, 0) }));
  } finally { await browser.close(); }
}
