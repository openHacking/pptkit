import assert from 'node:assert/strict';
import { chromium, firefox, webkit } from '@playwright/test';
import { createPresentation } from '../../packages/core/dist/index.js';
import { renderPresentationToSvg } from '../../packages/svg-renderer/dist/index.js';
const p = createPresentation(), slide = p.addSlide(), text = 'שלום (ABC 123) סוף';
for (const [i, direction] of ['ltr', 'rtl'].entries()) slide.addElement({
  type: 'text', id: direction, box: { x: 20, y: 20 + i * 100, width: 400, height: 80 }, frame: { margin: 0, wrap: true },
  content: [{ style: { direction, align: 'left' }, runs: [{ text, style: { fontFamily: 'Arial', fontSize: 24 } }] }],
});
const svg = (await renderPresentationToSvg(p)).slides[0].svg;
for (const [name, engine] of [['chromium', chromium], ['firefox', firefox], ['webkit', webkit]]) {
  const browser = await engine.launch({ headless: true });
  try {
    const page = await browser.newPage(); await page.setContent(svg);
    const items = await page.evaluate(() => ['ltr', 'rtl'].map(id => {
      const group = document.querySelector(`[data-pptkit-element-id="${id}"]`), paragraph = group.querySelector('div > div'), text = paragraph.querySelector('span[lang]').firstChild;
      const word = (start, end) => { const r = document.createRange(); r.setStart(text, start); r.setEnd(text, end); const box = r.getBoundingClientRect(); return { x: box.x, width: box.width }; };
      return { direction: getComputedStyle(paragraph).direction, alignment: getComputedStyle(paragraph).textAlign, first: word(0, 4), last: word(text.length - 3, text.length), text: text.data };
    }));
    assert.equal(items[0].text, text); assert.equal(items[1].text, text);
    assert.deepEqual(items.map(item => [item.direction, item.alignment]), [['ltr', 'left'], ['rtl', 'left']]);
    assert.ok(items[0].first.x < items[0].last.x, `${name}: LTR word order`);
    assert.ok(items[1].first.x > items[1].last.x, `${name}: RTL word order`);
    console.log(JSON.stringify({ browser: name, version: browser.version(), items }));
  } finally { await browser.close(); }
}
