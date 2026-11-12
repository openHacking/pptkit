import test from 'node:test';
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { readFile } from 'node:fs/promises';
import { importPptx } from '../packages/pptx-parser/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';
import { readGeneratedZip } from '../tests/support/read-generated-zip.mjs';

const hash = bytes => createHash('sha256').update(bytes).digest('hex');
const blue = 'fd298a04c9d501c6bc2f4900c902cbe49f5031717600c7a0de9a6ebe83803ee9';
const red = '7a7206f45f9ba5bd42b1f7d40889ad8ccc1bfe78e5bb6d442e8dd256d0d8a9b9';

test('native SVG resave changes projected fallback without changing the original-byte save contract', async () => {
  for (const [file, pairHash] of [['source', blue], ['resave', red]]) {
    const bytes = await readFile(`tests/fixtures/svg-pair-powerpoint-${file}.pptx`);
    const session = await importPptx(bytes, { projection: 'content' });
    try {
      const elements = session.document.slides[0].elements;
      assert.deepEqual(elements.map(e => e.name), ['SVG red / PNG blue', 'PNG blue control']);
      const assetFor = element => session.document.assets.find(a => a.id === element.assetId);
      assert.equal(hash(Buffer.from(assetFor(elements[0]).source.value, 'base64')), pairHash);
      assert.equal(hash(Buffer.from(assetFor(elements[1]).source.value, 'base64')), blue);
      const findings = session.report.findings.filter(f => f.reason.includes('An SVG image extension is present'));
      assert.ok(findings.some(f => f.operation === 'preview' && f.elementId === elements[0].id && f.tier === 'unsupported'));
      assert.ok(findings.some(f => f.operation === 'no-op-save' && f.elementId === elements[0].id && f.tier === 'preserved'));
      assert.ok(!findings.some(f => f.elementId === elements[1].id));
      const plan = await session.prepareSave({ mode: 'no-op' });
      assert.equal(plan.status, 'ready');
      assert.deepEqual((await session.save(plan.plan)).bytes, new Uint8Array(bytes));
      const rebuilt = readGeneratedZip((await generatePptx(session.document)).bytes);
      assert.ok(![...rebuilt.keys()].some(name => name.endsWith('.svg')));
      assert.doesNotMatch(rebuilt.get('ppt/slides/slide1.xml').toString(), /svgBlip/);
    } finally { session.dispose(); }
  }
});
