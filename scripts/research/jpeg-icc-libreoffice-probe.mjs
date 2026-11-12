// Optional independent consumer experiment; requires soffice, pdftoppm and Pillow.
import { mkdtemp, mkdir, readFile, rm, writeFile } from 'node:fs/promises';
import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import assert from 'node:assert/strict';
import { createPresentation } from '../../packages/core/dist/index.js';
import { generatePptx, inspectJpegIcc } from '../../packages/pptx-exporter/dist/index.js';
import { importPptx } from '../../packages/pptx-parser/dist/index.js';

const dir = await mkdtemp(join(tmpdir(), 'pptkit-jpeg-icc-'));
function run(command, args) {
  const result = spawnSync(command, args, { encoding: 'utf8', timeout: 60000 });
  assert.equal(result.status, 0, `${command}: ${result.error?.message ?? result.stderr}`);
  return result.stdout.trim();
}
function withoutIcc(source) {
  const pieces = [source.subarray(0, 2)]; let pos = 2;
  while (pos < source.length) {
    const start = pos; assert.equal(source[pos++], 255); while (source[pos] === 255) pos++;
    const marker = source[pos++];
    if (marker === 0xda || marker === 0xd9) { pieces.push(source.subarray(start)); break; }
    const length = source.readUInt16BE(pos), end = pos + length;
    assert.ok(length >= 2 && end <= source.length);
    if (!(marker === 0xe2 && source.subarray(pos + 2, pos + 14).equals(Buffer.from('ICC_PROFILE\0')))) pieces.push(source.subarray(start, end));
    pos = end;
  }
  return Buffer.concat(pieces);
}
try {
  const tagged = await readFile('tests/fixtures/jpeg-icc-srgb.jpg'), plain = withoutIcc(tagged);
  assert.equal(inspectJpegIcc(tagged).status, 'complete'); assert.equal(inspectJpegIcc(plain).status, 'no-profile');
  const deck = createPresentation({ size: { width: 100, height: 100 } });
  for (const [id, bytes] of [['tagged', tagged], ['plain', plain]]) deck.registerAsset({ id, kind: 'image', mimeType: 'image/jpeg', source: { type: 'bytes', value: bytes } });
  deck.addSlide({ elements: ['tagged', 'plain'].map((assetId, i) => ({ type: 'image', assetId, box: { x: 5 + i * 50, y: 5, width: 40, height: 40 } })) });
  const output = await generatePptx(deck, { deduplicateResolvedAssets: true, jpegIccDiagnostics: 'warn' });
  assert.deepEqual(output.warnings.map(w => [w.code, w.assetId]), [['jpeg-icc-color-unverified', 'tagged']]);
  let source = join(dir, 'input.pptx'); await writeFile(source, output.bytes);
  const rounds = [];
  for (let round = 0; round <= 3; round++) {
    const out = join(dir, `round-${round}`); await mkdir(out);
    if (round) { run('soffice', ['--headless', '--convert-to', 'pptx', '--outdir', out, source]); source = join(out, 'input.pptx'); }
    const inventory = JSON.parse(run('python3', ['-c', `import hashlib,io,json,posixpath,sys,zipfile,xml.etree.ElementTree as E
from PIL import Image
with zipfile.ZipFile(sys.argv[1]) as z:
    rels=E.fromstring(z.read('ppt/slides/_rels/slide1.xml.rels'))
    targets={r.get('Id'):posixpath.normpath(posixpath.join('ppt/slides',r.get('Target'))) for r in rels}
    slide=E.fromstring(z.read('ppt/slides/slide1.xml')); ns='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
    out=[]
    for b in slide.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}blip'):
        target=targets[b.get(ns+'embed')]; data=z.read(target); im=Image.open(io.BytesIO(data)); profile=im.info.get('icc_profile')
        out.append({'target':target,'sha256':hashlib.sha256(data).hexdigest(),'profileBytes':len(profile) if profile else 0,'profileSha256':hashlib.sha256(profile).hexdigest() if profile else None,'pixel':list(im.convert('RGB').getpixel((3,2)))})
    print(json.dumps(out))`, source]));
    assert.equal(inventory.length, 2);
    if (round === 0) {
      assert.notEqual(inventory[0].target, inventory[1].target);
      assert.equal(inventory[0].profileBytes, inspectJpegIcc(tagged).profileBytes);
      assert.equal(inventory[1].profileBytes, 0);
      assert.deepEqual(inventory[0].pixel, inventory[1].pixel);
    }
    run('soffice', ['--headless', '--convert-to', 'pdf', '--outdir', out, source]);
    run('pdftoppm', ['-png', '-singlefile', '-r', '72', join(out, 'input.pdf'), join(out, 'page')]);
    const pixels = JSON.parse(run('python3', ['-c', "from PIL import Image; import json,sys; im=Image.open(sys.argv[1]).convert('RGB'); print(json.dumps([list(im.getpixel((x,25))) for x in [25,75]]))", join(out, 'page.png')]));
    const bytes = await readFile(source), session = await importPptx(bytes, { projection: 'content' });
    assert.equal(session.document.slides[0].elements.length, 2);
    assert.deepEqual((await session.save((await session.prepareSave({ mode: 'no-op' })).plan)).bytes, new Uint8Array(bytes)); session.dispose();
    rounds.push({ round, inventory, pixels, parserNoOpExact: true });
  }
  const sha = bytes => createHash('sha256').update(bytes).digest('hex');
  console.log(JSON.stringify({ version: run('soffice', ['--version']), inputSha256: { tagged: sha(tagged), plain: sha(plain) }, rounds }, null, 2));
} finally { await rm(dir, { recursive: true, force: true }); }
