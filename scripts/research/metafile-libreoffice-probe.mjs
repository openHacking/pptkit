// Optional local experiment; requires soffice and Python Pillow. No network access.
import { mkdtemp, mkdir, readFile, rm, writeFile } from 'node:fs/promises';
import { spawnSync } from 'node:child_process';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import assert from 'node:assert/strict';
import { createPresentation } from '../../packages/core/dist/index.js';
import { generatePptx } from '../../packages/pptx-exporter/dist/index.js';
import { importPptx } from '../../packages/pptx-parser/dist/index.js';
import { readGeneratedZip } from '../../tests/support/read-generated-zip.mjs';
import { imageFixture, encodeImages, picture } from '../../packages/pptx-parser/test/image-fixture.mjs';
import { shape } from '../../packages/pptx-parser/test/content-fixture.mjs';

const dir = await mkdtemp(join(tmpdir(), 'pptkit-metafile-probe-'));
function run(command, args) {
  const result = spawnSync(command, args, { encoding: 'utf8', timeout: 60000 });
  assert.equal(result.status, 0, `${command}: ${result.error?.message ?? result.stderr}`);
  return result.stdout.trim();
}
try {
  const svg = join(dir, 'sample.svg');
  await writeFile(svg, '<svg xmlns="http://www.w3.org/2000/svg" width="160" height="90"><rect width="160" height="90" fill="white"/><rect x="10" y="10" width="50" height="30" fill="red"/><circle cx="125" cy="65" r="20" fill="blue"/></svg>');
  const reports = [];
  for (const format of ['emf', 'wmf']) {
    const output = join(dir, format); await mkdir(output);
    run('soffice', ['--headless', '--convert-to', format, '--outdir', output, svg]);
    const vector = await readFile(join(output, `sample.${format}`));
    const fixture = imageFixture(picture() + shape({ id: '3' }));
    fixture.parts['media/raster.bin'] = vector;
    fixture.mediaTypes['media/raster.bin'] = `image/${format}`;
    const source = encodeImages(fixture), session = await importPptx(source, { projection: 'content' });
    assert.equal(session.document.assets.length, 0);
    assert.deepEqual(session.document.slides[0].elements.map(e => e.type), ['shape']);
    const unsupported = session.report.findings.find(f => /MIME\/signature/.test(f.reason));
    assert.ok(unsupported, 'unsupported vector must be reported');
    const plan = await session.prepareSave({ mode: 'no-op' });
    assert.equal(plan.status, 'ready');
    assert.deepEqual((await session.save(plan.plan)).bytes, new Uint8Array(source));
    const rebuild = await session.prepareSave({ mode: 'rebuild' });
    assert.equal(rebuild.status, 'blocked');
    session.dispose();
    run('soffice', ['--headless', '--convert-to', 'png', '--outdir', output, join(output, `sample.${format}`)]);
    const pngPath = join(output, 'sample.png'), png = await readFile(pngPath);
    const pixels = JSON.parse(run('python3', ['-c', `import json,sys
from PIL import Image
im=Image.open(sys.argv[1]).convert('RGB')
out={'size':list(im.size)}
for name,test in [('red',lambda r,g,b:r>200 and g<80 and b<80),('blue',lambda r,g,b:b>200 and r<80 and g<80)]:
    pts=[(x,y) for y in range(im.height) for x in range(im.width) if test(*im.getpixel((x,y)))]
    out[name]=[min(x for x,y in pts),min(y for x,y in pts),max(x for x,y in pts),max(y for x,y in pts)] if pts else None
print(json.dumps(out))`, pngPath]));
    const deck = createPresentation();
    deck.registerAsset({ id: 'vector', kind: 'image', mimeType: `image/${format}`, source: { type: 'bytes', value: vector } });
    deck.addSlide().addElement({ type: 'image', assetId: 'vector', box: { x: 0, y: 0, width: 160, height: 90 } });
    const result = await generatePptx(deck, { resolveAsset: () => ({ bytes: png, mimeType: 'image/png', vectorRasterization: format }) });
    assert.deepEqual(result.warnings.map(w => w.code), ['asset-vector-rasterized']);
    assert.deepEqual(readGeneratedZip(result.bytes).get('ppt/media/vector.png'), png);
    reports.push({ format, vectorBytes: vector.length, pngBytes: png.length, importFinding: unsupported.reason, noOpExact: true, rebuild: rebuild.status, pixels });
  }
  console.log(JSON.stringify({ version: run('soffice', ['--version']), sourceSize: [160, 90], reports }, null, 2));
} finally {
  await rm(dir, { recursive: true, force: true });
}
