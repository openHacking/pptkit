import assert from 'node:assert/strict';
import { readFile, writeFile, readdir } from 'node:fs/promises';
import path from 'node:path';
import { createHash } from 'node:crypto';
import { chromium } from '@playwright/test';
import { normalizePresentation } from '../packages/core/dist/index.js';
import { importPptx } from '../packages/pptx-parser/dist/index.js';
const evidencePath = 'docs/testing/gif-powerpoint-reopen-evidence.json';
const evidenceBytes = await readFile(evidencePath), evidence = JSON.parse(evidenceBytes);
const digest = bytes => createHash('sha256').update(bytes).digest('hex');
const browser = await chromium.launch(), page = await browser.newPage(), requests = [];
await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
const records = [];
try {
  const url = text => 'data:text/javascript;base64,' + Buffer.from(text).toString('base64');
  const core = url(await readFile('packages/core/dist/browser/index.js', 'utf8'));
  const parser = url((await readFile('packages/pptx-parser/dist/browser/index.js', 'utf8')).replaceAll('"@pptkit/core"', JSON.stringify(core)));
  for (const name of ['core','pptx-parser']) await page.addScriptTag({ content: await readFile(`packages/${name}/dist/browser/global.js`, 'utf8') });
  const gif = await readFile(`${evidence.directory}/motion.gif`);
  for (const record of evidence.packages) {
    const bytes = await readFile(`${evidence.directory}/${record.file}`); assert.equal(digest(bytes), record.sha256);
    const session = await importPptx(bytes, { projection: 'content' });
    let expected;
    try {
      const plan = await session.prepareSave({ mode: 'no-op' }); assert.equal(plan.status, 'ready');
      expected = { normalized: normalizePresentation(session.document), report: session.report, bytes: [...(await session.save(plan.plan)).bytes] };
      assert.deepEqual(expected.bytes, [...bytes]);
      assert.equal(expected.normalized.assets.length, 1);
      assert.equal(expected.normalized.assets[0].mimeType, 'image/gif');
      assert.deepEqual(Buffer.from(expected.normalized.assets[0].source.value, 'base64'), gif);
      const images = expected.normalized.slides[0].elements.filter(e => e.type === 'image');
      assert.equal(images.length, 2); assert.ok(images.every(image => image.assetId === expected.normalized.assets[0].id));
    } finally { session.dispose(); }
    for (const format of ['esm','global']) {
      const actual = await page.evaluate(async ({ bytes, parser, core }) => {
        const api = parser ? Object.assign({}, await import(parser), await import(core)) : PPTKit;
        const session = await api.importPptx(Uint8Array.from(bytes), { projection: 'content' });
        try { const plan = await session.prepareSave({ mode: 'no-op' }); if (plan.status !== 'ready') throw Error('No-op blocked');
          return { normalized: api.normalizePresentation(session.document), report: session.report, bytes: [...(await session.save(plan.plan)).bytes] };
        } finally { session.dispose(); }
      }, { bytes: [...bytes], parser: format === 'esm' ? parser : null, core });
      assert.deepEqual(actual, expected);
    }
    records.push({ file: record.file, sha256: digest(bytes), exactNoOp: true, nodeEsmGlobalParity: true, sharedGifBytes: true });
  }
  assert.deepEqual(requests, []);
} finally { await browser.close(); }
const files = ['pnpm-lock.yaml', 'scripts/build-browser-bundle.mjs'];
async function walk(directory) {
  for (const entry of await readdir(directory, { withFileTypes: true })) {
    const file = path.join(directory, entry.name);
    if (entry.isDirectory()) await walk(file);
    else if (file.endsWith('.ts') || file.endsWith('.js')) files.push(file);
  }
}
for (const name of ['core','pptx-parser']) for (const dir of ['src','dist']) await walk(`packages/${name}/${dir}`);
const dependencies = await Promise.all(files.sort().map(async file => ({ file, sha256: digest(await readFile(file)) })));
await writeFile('docs/testing/gif-powerpoint-reopen-parser-evidence.json', JSON.stringify({
  scope: 'Four named packages, content import and exact no-op in Node/browser ESM/global; no general edit-save or native fidelity guarantee.',
  nativeEvidenceSha256: digest(evidenceBytes), scriptSha256: digest(await readFile('scripts/inspect-gif-native-reopen.mjs')), records, dependencies,
}, null, 2) + '\n');
console.log(JSON.stringify({ inputs: records.length, exactNoOps: records.length * 3 }));
