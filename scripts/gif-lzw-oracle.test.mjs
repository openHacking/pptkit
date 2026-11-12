// Optional independent encoder/decoder oracle; requires Python and Pillow.
import test from 'node:test';
import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { validateGifReplacement } from '../packages/pptx-parser/dist/import/validate-gif-replacement.js';

function pillowGif(width, height, frames, interlaced) {
  const script = String.raw`
import sys,io,random
from PIL import Image
w,h,n,interlace=map(int,sys.argv[1:]); palette=[v for i in range(256) for v in (i,(i*7)%256,(i*13)%256)]; rng=random.Random(42)
images=[]
for frame in range(n):
 im=Image.new('P',(w,h)); im.putpalette(palette); im.putdata([rng.randrange(256) for _ in range(w*h)]); images.append(im)
out=io.BytesIO(); images[0].save(out,format='GIF',save_all=n>1,append_images=images[1:],duration=80,loop=0,disposal=2,optimize=False,interlace=bool(interlace)); sys.stdout.buffer.write(out.getvalue())
`;
  const result = spawnSync('python3', ['-c', script, String(width), String(height), String(frames), interlaced ? '1' : '0'], { timeout: 15000, maxBuffer: 8 * 1024 * 1024 });
  assert.equal(result.status, 0, result.stderr.toString()); return result.stdout;
}

test('Pillow-generated GIF streams validate across LZW code widths, interlace and animation', async () => {
  for (const [width, height, frames, interlaced] of [[1, 1, 1, false], [11, 5, 1, false], [64, 64, 1, true], [256, 128, 1, true], [32, 24, 3, false]]) {
    const bytes = pillowGif(width, height, frames, interlaced), usage = { work: 0 };
    assert.equal(await validateGifReplacement(bytes, 10000000, 1000000, usage, async () => {}), true, `${width}x${height}, frames=${frames}, interlace=${interlaced}`);
    assert.equal(usage.decodedBytes, width * height * frames);
  }
});
