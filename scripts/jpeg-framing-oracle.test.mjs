import test from 'node:test';
import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { readFile } from 'node:fs/promises';
import { validateJpegReplacement } from '../packages/pptx-parser/dist/import/validate-jpeg-replacement.js';

test('Pillow baseline and progressive JPEG streams have accepted framing', async t => {
  const program = `import base64,io,json,random
from PIL import Image
r=random.Random(42)
out=[]
for size in [(1,1),(17,13),(128,64)]:
 im=Image.frombytes('RGB',size,bytes(r.randrange(256) for _ in range(size[0]*size[1]*3)))
 for progressive in (False,True):
  b=io.BytesIO(); im.save(b,format='JPEG',progressive=progressive); out.append(base64.b64encode(b.getvalue()).decode())
print(json.dumps(out))`;
  const result = spawnSync('python3', ['-c', program], { encoding: 'utf8' });
  if (result.error?.code === 'ENOENT' || /No module named 'PIL'/.test(result.stderr)) return t.skip('Python Pillow is unavailable');
  assert.equal(result.status, 0, result.stderr);
  for (const encoded of JSON.parse(result.stdout)) {
    const bytes = Buffer.from(encoded, 'base64');
    assert.equal(await validateJpegReplacement(bytes, 1000000, { work: 0 }, async () => {}), true);
    assert.equal(await validateJpegReplacement(bytes.subarray(0, -2), 1000000, { work: 0 }, async () => {}), false);
  }
});

test('libjpeg-turbo strict decode agrees on named progression controls', async t => {
  const available = spawnSync('djpeg', ['-version'], { encoding: 'utf8' });
  if (available.error?.code === 'ENOENT') return t.skip('djpeg is unavailable');
  assert.equal(available.status, 0, available.stderr);
  const names = ['jpeg-sequential-multiscan.jpg', 'jpeg-12bit-sequential.jpg', 'jpeg-12bit-progressive.jpg', 'jpeg-progressive-components.jpg'];
  const images = await Promise.all(names.map(name => readFile(new URL(`../tests/fixtures/${name}`, import.meta.url))));
  const decode = bytes => spawnSync('djpeg', ['-strict'], { input: bytes, maxBuffer: 1024 * 1024, timeout: 15000 });
  for (const [index, bytes] of images.entries()) {
    const result = decode(bytes);
    assert.equal(result.status, 0, `${names[index]}: ${result.stderr}`);
    assert.equal(await validateJpegReplacement(bytes, 1000000, { work: 0 }, async () => {}), true);
  }
  const progressive = images[3], first = progressive.indexOf(Buffer.from([0xff, 0xda]));
  const invalid = Buffer.from(progressive); invalid[first + 7 + 2 * invalid[first + 4]] = 0x21;
  const result = decode(invalid);
  assert.notEqual(result.status, 0); assert.match(result.stderr.toString(), /progression/i);
  assert.equal(await validateJpegReplacement(invalid, 1000000, { work: 0 }, async () => {}), false);
});
