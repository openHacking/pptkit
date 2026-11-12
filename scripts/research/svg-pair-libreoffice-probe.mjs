// Optional independent consumer experiment; requires soffice, pdftoppm and Python Pillow.
import { mkdtemp, mkdir, readFile, rm, writeFile } from 'node:fs/promises';
import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import assert from 'node:assert/strict';
import { createPresentation } from '../../packages/core/dist/index.js';
import { generatePptx } from '../../packages/pptx-exporter/dist/index.js';
import { importPptx } from '../../packages/pptx-parser/dist/index.js';
import { readGeneratedZip } from '../../tests/support/read-generated-zip.mjs';
import { zip } from '../../packages/pptx-parser/test/zip-fixture.mjs';

const dir = await mkdtemp(join(tmpdir(), 'pptkit-svg-pair-'));
function run(command, args) {
  const result = spawnSync(command, args, { encoding: 'utf8', timeout: 60000 });
  assert.equal(result.status, 0, `${command}: ${result.error?.message ?? result.stderr}`);
  return result.stdout.trim();
}
try {
  const pngPath = join(dir, 'blue.png');
  run('python3', ['-c', "from PIL import Image; import sys; Image.new('RGB',(20,20),(0,0,255)).save(sys.argv[1])", pngPath]);
  const png = await readFile(pngPath), deck = createPresentation({ size: { width: 100, height: 100 } });
  deck.registerAsset({ id: 'image', kind: 'image', mimeType: 'image/png', source: { type: 'bytes', value: png } });
  deck.addSlide().addElement({ type: 'image', assetId: 'image', box: { x: 0, y: 0, width: 100, height: 100 } });
  const base = (await generatePptx(deck)).bytes, parts = readGeneratedZip(base);
  const slide = parts.get('ppt/slides/slide1.xml').toString();
  const rels = parts.get('ppt/slides/_rels/slide1.xml.rels').toString();
  const types = parts.get('[Content_Types].xml').toString();
  const extension = '<a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rIdSvgProbe"/></a:ext></a:extLst>';
  assert.match(slide, /<a:blip r:embed="rId\d+"><\/a:blip>/);
  parts.set('ppt/slides/slide1.xml', Buffer.from(slide.replace('</a:blip>', `${extension}</a:blip>`)));
  parts.set('ppt/slides/_rels/slide1.xml.rels', Buffer.from(rels.replace('</Relationships>', '<Relationship Id="rIdSvgProbe" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/probe.svg"/></Relationships>')));
  parts.set('[Content_Types].xml', Buffer.from(types.replace('</Types>', '<Default Extension="svg" ContentType="image/svg+xml"/></Types>')));
  const svg = Buffer.from('<svg xmlns="http://www.w3.org/2000/svg" width="20" height="20"><rect width="20" height="20" fill="red"/></svg>');
  parts.set('ppt/media/probe.svg', svg);
  const paired = zip([...parts].map(([name, body]) => ({ name, body }))).bytes;
  const imported = await importPptx(paired, { projection: 'content' });
  assert.equal(imported.document.assets.length, 1);
  assert.equal(imported.document.assets[0].source.value, png.toString('base64'));
  assert.ok(imported.report.findings.some(f => /An SVG image extension is present/.test(f.reason)));
  assert.deepEqual((await imported.save((await imported.prepareSave({ mode: 'no-op' })).plan)).bytes, new Uint8Array(paired));
  imported.dispose();
  const authored = createPresentation({ size: { width: 100, height: 100 } });
  authored.registerAsset({ id: 'image', kind: 'image', mimeType: 'image/svg+xml', source: { type: 'bytes', value: svg } });
  authored.addSlide().addElement({ type: 'image', assetId: 'image', box: { x: 0, y: 0, width: 100, height: 100 } });
  const generated = await generatePptx(authored, { resolveAsset: () => ({ bytes: png, mimeType: 'image/png', svgPair: svg }) });
  assert.deepEqual(generated.warnings.map(w => w.code), ['asset-svg-pair-unverified']);
  const results = [];
  for (const [name, bytes] of [['fallback', base], ['paired', paired], ['generated', generated.bytes]]) {
    const input = join(dir, `${name}.pptx`), out = join(dir, name); await mkdir(out); await writeFile(input, bytes);
    run('soffice', ['--headless', '--convert-to', 'pdf', '--outdir', out, input]);
    run('pdftoppm', ['-png', '-singlefile', '-r', '72', join(out, `${name}.pdf`), join(out, 'page')]);
    const pixel = JSON.parse(run('python3', ['-c', "from PIL import Image; import json,sys; im=Image.open(sys.argv[1]).convert('RGB'); print(json.dumps({'size':list(im.size),'center':list(im.getpixel((im.width//2,im.height//2)))}))", join(out, 'page.png')]));
    results.push({ name, ...pixel });
  }
  const rounds = [];
  let source = join(dir, 'generated.pptx');
  for (let round = 1; round <= 3; round++) {
    const out = join(dir, `round-${round}`); await mkdir(out);
    run('soffice', ['--headless', '--convert-to', 'pptx', '--outdir', out, source]);
    source = join(out, 'generated.pptx');
    const inventory = JSON.parse(run('python3', ['-c', `import hashlib,io,json,posixpath,sys,zipfile,xml.etree.ElementTree as E
from PIL import Image
with zipfile.ZipFile(sys.argv[1]) as z:
    slide=E.fromstring(z.read('ppt/slides/slide1.xml'))
    svg='{http://schemas.microsoft.com/office/drawing/2016/SVG/main}svgBlip'
    relns='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
    rels=E.fromstring(z.read('ppt/slides/_rels/slide1.xml.rels'))
    targets={r.get('Id'):posixpath.normpath(posixpath.join('ppt/slides',r.get('Target'))) for r in rels}
    base=[targets.get(x.get(relns+'embed')) for x in slide.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}blip')]
    vectors=[targets.get(x.get(relns+'embed')) for x in slide.iter(svg)]
    images=[]
    for n in sorted(z.namelist()):
        if not n.startswith('ppt/media/'): continue
        data=z.read(n); entry={'name':n,'sha256':hashlib.sha256(data).hexdigest()}
        if n.endswith('.png'):
            im=Image.open(io.BytesIO(data)).convert('RGB'); entry['size']=list(im.size); entry['center']=list(im.getpixel((im.width//2,im.height//2)))
        images.append(entry)
    print(json.dumps({'svgBlips':len(vectors),'baseTargets':base,'svgTargets':vectors,'images':images}))`, source]));
    run('soffice', ['--headless', '--convert-to', 'pdf', '--outdir', out, source]);
    run('pdftoppm', ['-png', '-singlefile', '-r', '72', join(out, 'generated.pdf'), join(out, 'page')]);
    const pixel = JSON.parse(run('python3', ['-c', "from PIL import Image; import json,sys; im=Image.open(sys.argv[1]).convert('RGB'); print(json.dumps(list(im.getpixel((im.width//2,im.height//2)))))", join(out, 'page.png')]));
    const bytes = await readFile(source), session = await importPptx(bytes, { projection: 'content' });
    const plan = await session.prepareSave({ mode: 'no-op' });
    assert.equal(plan.status, 'ready'); assert.deepEqual((await session.save(plan.plan)).bytes, new Uint8Array(bytes));
    session.dispose(); rounds.push({ round, ...inventory, center: pixel, parserNoOpExact: true });
  }
  const sha256 = bytes => createHash('sha256').update(bytes).digest('hex');
  console.log(JSON.stringify({ version: run('soffice', ['--version']), inputSha256: { png: sha256(png), svg: sha256(svg) }, results, rounds }, null, 2));
} finally { await rm(dir, { recursive: true, force: true }); }
