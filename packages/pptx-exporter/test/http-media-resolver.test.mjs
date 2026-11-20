import test from 'node:test';
import assert from 'node:assert/strict';
import { createServer } from 'node:http';
import { once } from 'node:events';
import { readFile } from 'node:fs/promises';
import { gzipSync } from 'node:zlib';
import { createPresentation } from '@pptkit/core';
import { createHttpMediaResolver, generatePptx } from '../dist/node.js';
import { readGeneratedZip } from '../../../tests/support/read-generated-zip.mjs';

test('HTTP media adapter authorizes redirects, bounds real streams, cancels and retries', async t => {
  const fixture = readGeneratedZip(await readFile(new URL('../../../tests/fixtures/pcm-wave-powerpoint.pptx', import.meta.url)));
  const wave = fixture.get('ppt/media/media1.wav'); let slowStarted;
  const slow = new Promise(resolve => { slowStarted = resolve; });
  const hits = [];
  const server = createServer((req, res) => {
    hits.push(req.url);
    if (req.url === '/redirect') { res.writeHead(302, { location: '/chunked' }); res.end(); return; }
    if (req.url === '/blocked') { res.writeHead(302, { location: '/forbidden' }); res.end(); return; }
    if (req.url === '/loop') { res.writeHead(302, { location: '/loop' }); res.end(); return; }
    if (req.url === '/error') { res.writeHead(503); res.end(); return; }
    if (req.url === '/slow') { res.writeHead(200, { 'content-type': 'audio/wav' }); res.write(wave.subarray(0, 16)); slowStarted(); return; }
    if (req.url === '/gzip') { res.writeHead(200, { 'content-type': 'audio/wav', 'content-encoding': 'gzip' }); res.end(gzipSync(wave)); return; }
    if (req.url === '/wrong') { res.writeHead(200, { 'content-type': 'video/mp4' }); res.end(wave); return; }
    res.writeHead(200, { 'content-type': 'audio/wav' }); res.write(wave.subarray(0, 100)); res.end(wave.subarray(100));
  });
  server.listen(0, '127.0.0.1'); await once(server, 'listening');
  t.after(() => { server.closeAllConnections(); server.close(); });
  const base = `http://127.0.0.1:${server.address().port}`;
  const make = path => {
    const p = createPresentation(); p.registerAsset({ id: 'sound', kind: 'audio', mimeType: 'audio/wav', source: { type: 'url', value: base + path } });
    p.addSlide({ elements: [{ type: 'audio', audioAssetId: 'sound', box: { x: 0, y: 0, width: 40, height: 40 } }] }); return p;
  };
  const authorized = [], resolver = createHttpMediaResolver({ maxRedirects: 1, authorizeUrl: url => { authorized.push(url); return url !== base + '/forbidden'; } });
  const options = { audioMode: 'wave-pcm-v1', resolveMedia: resolver, maxAssetBytes: wave.length };
  const output = await generatePptx(make('/redirect'), options);
  assert.deepEqual(readGeneratedZip(output.bytes).get('ppt/media/sound.wav'), wave);
  assert.deepEqual(authorized, [base + '/redirect', base + '/chunked']);
  await assert.rejects(generatePptx(make('/blocked'), options), /not authorized/); assert.ok(!hits.includes('/forbidden'));
  await assert.rejects(generatePptx(make('/loop'), options), /maxRedirects/);
  await assert.rejects(generatePptx(make('/error'), options), /503/);
  await assert.rejects(generatePptx(make('/wrong'), options), /incompatible MIME/);
  await assert.rejects(generatePptx(make('/chunked'), { ...options, maxAssetBytes: wave.length - 1 }), /maxBytes/);
  await assert.rejects(generatePptx(make('/gzip'), { ...options, maxAssetBytes: wave.length - 1 }), /maxBytes/);
  const controller = new AbortController(), reason = new Error('cancel stream');
  const pending = generatePptx(make('/slow'), { ...options, signal: controller.signal });
  await slow; controller.abort(reason); await assert.rejects(pending, error => error === reason);
  assert.deepEqual(readGeneratedZip((await generatePptx(make('/chunked'), options)).bytes).get('ppt/media/sound.wav'), wave);
});

test('HTTP media adapter requires explicit valid authorization and redirect limits', () => {
  assert.throws(() => createHttpMediaResolver({}), /authorizeUrl/);
  for (const maxRedirects of [-1, 1.5, 11]) assert.throws(() => createHttpMediaResolver({ authorizeUrl: () => true, maxRedirects }), /maxRedirects/);
});

test('invalid URLs, path sources and invalid budgets fail before authorization', async () => {
  let calls = 0;
  const resolver = createHttpMediaResolver({ authorizeUrl() { calls++; return true; } });
  for (const value of ['file:///tmp/test.wav', 'http://user:pass@example.test/x', 'https://example.test/x#part']) {
    await assert.rejects(resolver({ source: { type: 'url', value } }, { maxBytes: 10 }), /Media URL/);
  }
  await assert.rejects(resolver({ source: { type: 'path', value: '/tmp/x' } }, { maxBytes: 10 }), /URL source/);
  await assert.rejects(resolver({ source: { type: 'url', value: 'https://example.test' } }, { maxBytes: NaN }), /maxBytes/);
  assert.equal(calls, 0);
});
