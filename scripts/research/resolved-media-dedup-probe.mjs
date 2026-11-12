// Optional independent consumer experiment; requires soffice, pdftoppm and Pillow.
import { mkdtemp, mkdir, readFile, rm, writeFile } from 'node:fs/promises';
import { spawnSync } from 'node:child_process';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import assert from 'node:assert/strict';
import { createPresentation } from '../../packages/core/dist/index.js';
import { generatePptx } from '../../packages/pptx-exporter/dist/index.js';
import { importPptx } from '../../packages/pptx-parser/dist/index.js';

const dir = await mkdtemp(join(tmpdir(), 'pptkit-dedup-probe-'));
function run(command, args) {
  const result = spawnSync(command, args, { encoding: 'utf8', timeout: 60000 });
  assert.equal(result.status, 0, `${command}: ${result.error?.message ?? result.stderr}`);
  return result.stdout.trim();
}
try {
  const redPath = join(dir, 'red.png'), bluePath = join(dir, 'blue.png');
  run('python3', ['-c', "from PIL import Image; import sys; Image.new('RGB',(20,20),(255,0,0)).save(sys.argv[1]); Image.new('RGB',(20,20),(0,0,255)).save(sys.argv[2])", redPath, bluePath]);
  const red = await readFile(redPath), blue = await readFile(bluePath), deck = createPresentation({ size: { width: 100, height: 100 } });
  deck.registerAsset({ id: 'red-bytes', kind: 'image', mimeType: 'image/png', source: { type: 'bytes', value: red } });
  deck.registerAsset({ id: 'red-path', kind: 'image', mimeType: 'image/png', source: { type: 'path', value: redPath } });
  deck.registerAsset({ id: 'blue', kind: 'image', mimeType: 'image/png', source: { type: 'bytes', value: blue } });
  const image = (id, assetId, x, y) => ({ id, type: 'image', assetId, box: { x, y, width: 20, height: 20 } });
  deck.defineSlideMaster({ id: 'master', name: 'Master', elements: [image('master-red', 'red-bytes', 5, 70)] });
  deck.defineSlideLayout({ id: 'layout', name: 'Layout', masterId: 'master', elements: [image('layout-red', 'red-path', 35, 70)] });
  deck.addSlide({ layoutId: 'layout', elements: [image('slide1-red', 'red-bytes', 5, 10), image('slide1-red2', 'red-path', 35, 10)] });
  deck.addSlide({ layoutId: 'layout', elements: [image('slide2-red', 'red-path', 5, 10), image('slide2-blue', 'blue', 65, 10)] });
  const calls = [], generated = await generatePptx(deck, { deduplicateResolvedAssets: true, resolveAsset: asset => { calls.push(asset.id); return { bytes: asset.id === 'blue' ? blue : red }; } });
  assert.deepEqual(calls, ['red-bytes', 'red-path', 'blue']); assert.deepEqual(generated.warnings, []);
  let source = join(dir, 'input.pptx'); await writeFile(source, generated.bytes);
  const rounds = [];
  for (let round = 0; round <= 3; round++) {
    const out = join(dir, `round-${round}`); await mkdir(out);
    if (round > 0) { run('soffice', ['--headless', '--convert-to', 'pptx', '--outdir', out, source]); source = join(out, 'input.pptx'); }
    const inventory = JSON.parse(run('python3', ['-c', `import hashlib,io,json,posixpath,sys,zipfile,xml.etree.ElementTree as E
from PIL import Image
with zipfile.ZipFile(sys.argv[1]) as z:
    result={}
    for owner in ['ppt/slides/slide1.xml','ppt/slides/slide2.xml','ppt/slideLayouts/slideLayout2.xml','ppt/slideMasters/slideMaster2.xml']:
        folder,name=posixpath.split(owner); relpath=folder+'/_rels/'+name+'.rels'
        rels=E.fromstring(z.read(relpath)); targets={r.get('Id'):posixpath.normpath(posixpath.join(folder,r.get('Target'))) for r in rels}
        root=E.fromstring(z.read(owner)); ns='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
        values=[]
        for blip in root.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}blip'):
            target=targets.get(blip.get(ns+'embed'))
            if not target: continue
            data=z.read(target); im=Image.open(io.BytesIO(data)).convert('RGB')
            values.append({'target':target,'sha256':hashlib.sha256(data).hexdigest(),'center':list(im.getpixel((im.width//2,im.height//2)))})
        result[owner]=values
    print(json.dumps(result))`, source]));
    run('soffice', ['--headless', '--convert-to', 'pdf', '--outdir', out, source]);
    run('pdftoppm', ['-png', '-r', '72', join(out, 'input.pdf'), join(out, 'page')]);
    const pixels = JSON.parse(run('python3', ['-c', `from PIL import Image
import json,sys
out=[]
for p in sys.argv[1:]:
    im=Image.open(p).convert('RGB'); out.append([list(im.getpixel((x,y))) for x,y in [(15,20),(45,20),(75,20),(15,80),(45,80)]])
print(json.dumps(out))`, join(out, 'page-1.png'), join(out, 'page-2.png')]));
    assert.deepEqual(pixels, [
      [[255,0,0],[255,0,0],[255,255,255],[255,0,0],[255,0,0]],
      [[255,0,0],[255,255,255],[0,0,255],[255,0,0],[255,0,0]],
    ]);
    const first = inventory['ppt/slides/slide1.xml'], second = inventory['ppt/slides/slide2.xml'];
    const shared = [...inventory['ppt/slideLayouts/slideLayout2.xml'], ...inventory['ppt/slideMasters/slideMaster2.xml']];
    assert.equal(first.length, 2); assert.equal(second.length, 2); assert.equal(shared.length, 2);
    const redHash = first[0].sha256;
    assert.ok([first[1], second[0], ...shared].every(item => item.sha256 === redHash && item.target === first[0].target));
    assert.deepEqual(second[1].center, [0,0,255]); assert.notEqual(second[1].target, first[0].target);
    if (round === 0) { assert.equal(inventory['ppt/slideLayouts/slideLayout2.xml'].length, 1); assert.equal(inventory['ppt/slideMasters/slideMaster2.xml'].length, 1); }
    const bytes = await readFile(source), session = await importPptx(bytes, { projection: 'content' });
    assert.deepEqual((await session.save((await session.prepareSave({ mode: 'no-op' })).plan)).bytes, new Uint8Array(bytes)); session.dispose();
    rounds.push({ round, inventory, pixels, parserNoOpExact: true });
  }
  console.log(JSON.stringify({ version: run('soffice', ['--version']), rounds }, null, 2));
} finally { await rm(dir, { recursive: true, force: true }); }
