import test from 'node:test';
import assert from 'node:assert/strict';
import { createServer } from 'node:http';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { chromium } from '@playwright/test';
import { importPptx } from '../packages/pptx-parser/dist/index.js';

test('React example updates isolated preview and downloads the current title', async () => {
  const dist = fileURLToPath(new URL('../examples/react-preview/dist/', import.meta.url));
  const server = createServer(async (req, res) => {
    const pathname = new URL(req.url, 'http://localhost').pathname;
    if (pathname !== '/' && !/^\/assets\/[a-zA-Z0-9_.-]+\.js$/.test(pathname)) { res.writeHead(404).end(); return; }
    try {
      const data = await readFile(path.join(dist, pathname === '/' ? 'index.html' : pathname));
      res.writeHead(200, { 'content-type': pathname === '/' ? 'text/html' : 'text/javascript' }).end(data);
    } catch { res.writeHead(404).end(); }
  });
  await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
  const browser = await chromium.launch({ headless: true });
  try {
    const page = await browser.newPage();
    const errors = [], external = [];
    page.on('pageerror', error => errors.push(error.message));
    page.on('request', request => { if (!request.url().startsWith(`http://127.0.0.1:${server.address().port}/`)) external.push(request.url()); });
    await page.goto(`http://127.0.0.1:${server.address().port}/`);
    await page.getByRole('status').getByText('Preview ready').waitFor();
    const frame = page.frameLocator('iframe[title="Slide preview"]');
    await frame.getByText('Quarterly review').waitFor();
    const detail = page.frameLocator('iframe[title="Slide preview 2"]');
    await detail.getByText('Supporting details').waitFor();
    const detailBefore = await page.locator('iframe[title="Slide preview 2"]').getAttribute('srcdoc');
    const title = 'Updated <review> & plan';
    await page.getByLabel('Slide title').fill(title);
    await frame.getByText(title).waitFor();
    assert.equal(await page.locator('iframe[title="Slide preview"]').getAttribute('sandbox'), '');
    assert.equal(await page.locator('iframe[title="Slide preview 2"]').getAttribute('srcdoc'), detailBefore);
    for (const value of ['Rapid one', 'Rapid two', title]) await page.getByLabel('Slide title').fill(value);
    await frame.getByText(title).waitFor();
    const [download] = await Promise.all([page.waitForEvent('download'), page.getByRole('button', { name: 'Download PPTX' }).click()]);
    assert.equal(download.suggestedFilename(), 'review.pptx');
    const stream = await download.createReadStream();
    const chunks = []; for await (const chunk of stream) chunks.push(chunk);
    const session = await importPptx(Buffer.concat(chunks), { projection: 'content' });
    try {
      const text = session.document.slides[0].elements.find(element => element.type === 'text');
      assert.equal(text.content[0].runs[0].text, title);
      assert.equal(session.document.slides.length, 2);
      const detailText = session.document.slides[1].elements.find(element => element.type === 'text');
      assert.equal(detailText.content[0].runs[0].text, 'Supporting details');
    } finally { session.dispose(); }
    assert.deepEqual(errors, []);
    assert.deepEqual(external, []);
  } finally {
    await browser.close();
    await new Promise((resolve, reject) => server.close(error => error ? reject(error) : resolve()));
  }
});
