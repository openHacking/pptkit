import assert from 'node:assert/strict';
import test from 'node:test';
import { chromium } from '@playwright/test';
import { previewHtml } from '../packages/cli/dist/commands/preview-html.js';

test('HTML preview isolates SVG IDs and disables scripts in each slide frame', async () => {
  const svg = color => `<svg xmlns="http://www.w3.org/2000/svg" width="100" height="50"><defs><linearGradient id="same"><stop stop-color="${color}"/></linearGradient></defs><rect width="100" height="50" fill="url(#same)"/><script>document.body.dataset.executed='yes'</script></svg>`;
  const html = previewHtml([{ index: 0, hidden: false, svg: svg('red') }, { index: 1, hidden: true, svg: svg('blue') }], 100, 50, 100000);
  const browser = await chromium.launch({ headless: true });
  try {
    const page = await browser.newPage(); await page.setContent(html);
    for (const [index, color] of ['red', 'blue'].entries()) {
      const frame = page.frameLocator('iframe').nth(index);
      assert.equal(await frame.locator('linearGradient stop').getAttribute('stop-color'), color);
      assert.equal(await frame.locator('rect').getAttribute('fill'), 'url(#same)');
      assert.equal(await frame.locator('body').getAttribute('data-executed'), null);
      assert.equal(await frame.locator('svg').count(), 1);
    }
    assert.equal(await page.locator('section[data-hidden="true"]').count(), 1);
  } finally { await browser.close(); }
});
