// Optional producer experiment. Headless conversion is not an interactive editability claim.
import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { createHash } from 'node:crypto';
import { mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises';
import { arch, platform, release } from 'node:os';
import path from 'node:path';
import { pathToFileURL } from 'node:url';

const run = promisify(execFile), executable = process.argv[2] ?? 'soffice';
const options = { timeout: 60000, maxBuffer: 8 * 1024 * 1024 };
const application = (await run(executable, ['--version'], options)).stdout.trim();
await run(process.execPath, ['scripts/prepare-cross-copy-native.mjs'], options);
await mkdir('.tmp/cross-copy-libreoffice', { recursive: true });
const directory = await mkdtemp(path.resolve('.tmp/cross-copy-libreoffice/run-'));
const initial = await readFile('.tmp/cross-document-copy-native/input.pptx');
let current = path.join(directory, 'input.pptx'); await writeFile(current, initial);
const profile = `-env:UserInstallation=${pathToFileURL(path.join(directory, 'profile')).href}`;
const files = [current], logs = [];
for (let round = 1; round <= 3; round++) {
  const outdir = path.join(directory, `round-${round}`); await mkdir(outdir);
  const result = await run(executable, [profile, '--headless', '--convert-to', 'pptx:Impress Office Open XML', '--outdir', outdir, current], options);
  logs.push({ stdout: result.stdout, stderr: result.stderr }); current = path.join(outdir, 'input.pptx');
  await readFile(current); files.push(current);
}
const pdfDir = path.join(directory, 'pdf'); await mkdir(pdfDir);
const pdfLog = await run(executable, [profile, '--headless', '--convert-to', 'pdf:impress_pdf_Export', '--outdir', pdfDir, current], options);
const pdf = path.join(pdfDir, 'input.pdf'); await readFile(pdf);
const python = String.raw`
import sys,json,hashlib,zipfile,xml.etree.ElementTree as E,fitz,re,posixpath
P='http://schemas.openxmlformats.org/presentationml/2006/main'
A='http://schemas.openxmlformats.org/drawingml/2006/main'
R='http://schemas.openxmlformats.org/package/2006/relationships'
O='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
def inspect(file):
  with zipfile.ZipFile(file) as z:
    names=z.namelist()
    slides=sorted((n for n in names if re.fullmatch(r'ppt/slides/slide\d+\.xml',n)),key=lambda n:int(re.search(r'\d+',n).group()))
    notes=sorted((n for n in names if re.fullmatch(r'ppt/notesSlides/notesSlide\d+\.xml',n)),key=lambda n:int(re.search(r'\d+',n).group()))
    def text(name): return ''.join(n.text or '' for n in E.fromstring(z.read(name)).iter('{%s}t'%A))
    def links(name):
      root=E.fromstring(z.read(name))
      return [{'id':x.get('Id'),'type':x.get('Type','').split('/')[-1],'target':x.get('Target')} for x in root if x.tag=='{%s}Relationship'%R]
    noteLinks={n:links(n) for n in names if re.fullmatch(r'ppt/notesSlides/_rels/notesSlide\d+\.xml\.rels',n)}
    slideLinks={n:links(n) for n in names if re.fullmatch(r'ppt/slides/_rels/slide\d+\.xml\.rels',n)}
    root=E.fromstring(z.read('ppt/slides/slide2.xml')); objects=[]
    for node in root.find('.//{%s}spTree'%P):
      kind=node.tag.split('}')[-1]
      if kind not in ('sp','pic','cxnSp'): continue
      properties=node.find('.//{%s}cNvPr'%P); transform=node.find('{%s}spPr/{%s}xfrm'%(P,A))
      offset=transform.find('{%s}off'%A) if transform is not None else None; extent=transform.find('{%s}ext'%A) if transform is not None else None
      click=properties.find('{%s}hlinkClick'%A) if properties is not None else None
      objects.append({'kind':kind,'name':properties.get('name') if properties is not None else None,'offset':dict(offset.attrib) if offset is not None else None,'extent':dict(extent.attrib) if extent is not None else None,'clickId':click.get('{%s}id'%O) if click is not None else None,'customGeometry':node.find('.//{%s}custGeom'%A) is not None})
    return {'slideText':[text(n) for n in slides], 'notesText':[text(n) for n in notes], 'notesRelationships':noteLinks, 'slideRelationships':slideLinks, 'copiedPageObjects':objects,
      'media':[{'name':n,'sha256':hashlib.sha256(z.read(n)).hexdigest()} for n in names if n.startswith('ppt/media/') and not n.endswith('/')]}
items=[inspect(x) for x in sys.argv[1:-1]]
assert len(items[0]['slideText'])==2 and 'Control page' in items[0]['slideText'][0] and 'Copied page' in items[0]['slideText'][1]
for item in items:
  assert len(item['slideText'])==2 and 'Control page' in item['slideText'][0] and 'Copied page' in item['slideText'][1]
  assert len(item['notesText'])==2 and 'Control page notes' in item['notesText'][0] and 'Copied page notes' in item['notesText'][1]
  for i in (1,2):
    assert any(r['type']=='slide' and r['target'].endswith('/slide%d.xml'%i) for r in item['notesRelationships']['ppt/notesSlides/_rels/notesSlide%d.xml.rels'%i])
    assert any(r['type']=='notesSlide' and r['target'].endswith('/notesSlide%d.xml'%i) for r in item['slideRelationships']['ppt/slides/_rels/slide%d.xml.rels'%i])
  assert len(item['media'])==1 and item['media'][0]['sha256']==items[0]['media'][0]['sha256']
  objects={o['name']:o for o in item['copiedPageObjects']}
  assert set(objects)=={'element-1','element-2','element-3','element-4'} and objects['element-2']['kind']=='pic' and objects['element-4']['customGeometry']
  assert any(r['id']==objects['element-1']['clickId'] and r['type']=='slide' and posixpath.normpath(posixpath.join('ppt/slides',r['target']))=='ppt/slides/slide2.xml' for r in item['slideRelationships']['ppt/slides/_rels/slide2.xml.rels'])
document=fitz.open(sys.argv[-1]); pages=[p.get_text() for p in document]
assert len(pages)==2 and 'Control page' in pages[0] and 'Copied page' in pages[1]
initial={o['name']:o for o in items[0]['copiedPageObjects']}; final={o['name']:o for o in items[-1]['copiedPageObjects']}
deltas={name:{part:{key:int(value)-int(initial[name][part][key]) for key,value in final[name][part].items()} for part in ('offset','extent')} for name in initial}
print(json.dumps({'oracle':'Python zipfile, ElementTree and PyMuPDF','rounds':items,'finalGeometryDeltasEmu':deltas,'reopenedPdfPageText':pages}))
`;
const independent = JSON.parse((await run('python3', ['-c', python, ...files, pdf], options)).stdout);
const hash = value => createHash('sha256').update(value).digest('hex');
const fileHashes = await Promise.all(files.map(async file => ({ file, sha256: hash(await readFile(file)) })));
const evidence = { schemaVersion: 1, scope: 'Two-page copied-slide fixture; LibreOfficeDev headless three successive saves and final PDF reopening. No interactive edit, repair-dialog, PowerPoint or full visual-fidelity claim.', environment: { application, platform: platform(), release: release(), arch: arch() }, directory, fileHashes, pdf: { file: pdf, sha256: hash(await readFile(pdf)) }, independent, logs, pdfLog };
assert.equal(independent.rounds.length, 4);
await writeFile('docs/testing/cross-copy-libreoffice-evidence.json', JSON.stringify(evidence, null, 2) + '\n');
console.log(JSON.stringify({ directory, application, rounds: independent.rounds.map(item => ({ slideText: item.slideText, notesText: item.notesText, media: item.media.length })), pdfPages: independent.reopenedPdfPageText.length }, null, 2));
