// Independent OOXML and headless producer observation for strict cross-document placeholder indexes.
import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { createHash } from 'node:crypto';
import { mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises';
import { arch, platform, release } from 'node:os';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { createPresentation, validatePresentation } from '../packages/core/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';

const run = promisify(execFile), executable = process.argv[2] ?? 'soffice';
const opts = { timeout: 60000, maxBuffer: 8 * 1024 * 1024 };
const application = (await run(executable, ['--version'], opts)).stdout.trim();
const title = { x: 60, y: 40, width: 400, height: 60 }, body = { x: 60, y: 130, width: 400, height: 60 };
const source = createPresentation(), dest = createPresentation();
source.defineSlideLayout({ id: 'source', name: 'Source', placeholders: [{ key: 'heading', kind: 'title', index: 0, box: title }, { key: 'body', kind: 'body', index: 4294967295, box: body }] });
for (const id of ['a', 'b']) source.addSlide({ id, layoutId: 'source', notes: `Notes ${id.toUpperCase()}`, elements: [
  { id: `title-${id}`, type: 'text', placeholderKey: 'heading', content: `Title ${id.toUpperCase()}` },
  { id: `body-${id}`, type: 'text', placeholderKey: 'body', content: `Body ${id.toUpperCase()}` },
] });
for (const id of ['a', 'b']) dest.defineSlideLayout({ id: `layout-${id}`, name: `Layout ${id}`, placeholders: [
  { key: `heading-${id}`, kind: 'title', index: 0, box: title }, { key: `body-${id}`, kind: 'body', index: 4294967295, box: body },
] });
const copied = dest.copySlidesFrom(source, ['a', 'b'], { layoutId: null, layouts: { a: 'layout-a', b: 'layout-b' }, theme: 'destination', geometry: 'preserve', customData: 'reject', placeholderRoles: 'require-match-and-index', assets: { deduplicate: 'content', external: 'reject' }, perSlide: {
  a: { placeholders: { heading: 'heading-a', body: 'body-a' } }, b: { placeholders: { heading: 'heading-b', body: 'body-b' } },
} });
assert.equal(copied.slides.length, 2); assert.deepEqual(validatePresentation(dest), []);
source.removeSlide('a'); source.removeSlide('b');
await mkdir('.tmp/copy-placeholder-index-libreoffice', { recursive: true });
const directory = await mkdtemp(path.resolve('.tmp/copy-placeholder-index-libreoffice/run-'));
let current = path.join(directory, 'input.pptx'); await writeFile(current, (await generatePptx(dest)).bytes);
const files = [current], logs = [], profile = `-env:UserInstallation=${pathToFileURL(path.join(directory, 'profile')).href}`;
for (let round = 1; round <= 3; round++) {
  const outdir = path.join(directory, `round-${round}`); await mkdir(outdir);
  const result = await run(executable, [profile, '--headless', '--convert-to', 'pptx:Impress Office Open XML', '--outdir', outdir, current], opts);
  logs.push({ stdout: result.stdout, stderr: result.stderr }); current = path.join(outdir, 'input.pptx'); await readFile(current); files.push(current);
}
const pdfDir = path.join(directory, 'pdf'); await mkdir(pdfDir);
const pdfLog = await run(executable, [profile, '--headless', '--convert-to', 'pdf:impress_pdf_Export', '--outdir', pdfDir, current], opts);
const pdf = path.join(pdfDir, 'input.pdf'); await readFile(pdf);
const python = String.raw`
import sys,json,zipfile,xml.etree.ElementTree as E,fitz,posixpath
P='http://schemas.openxmlformats.org/presentationml/2006/main'; A='http://schemas.openxmlformats.org/drawingml/2006/main'; R='http://schemas.openxmlformats.org/package/2006/relationships'
def inspect(file):
  with zipfile.ZipFile(file) as z:
    def rels(part):
      d,n=posixpath.split(part); name=f'{d}/_rels/{n}.rels'
      if name not in z.namelist(): return []
      return [{'type':r.get('Type','').split('/')[-1], 'mode':r.get('TargetMode','Internal'), 'target':posixpath.normpath(posixpath.join(d,r.get('Target')))} for r in E.fromstring(z.read(name)) if r.tag=='{%s}Relationship'%R]
    def one(part,kind):
      xs=[r['target'] for r in rels(part) if r['type']==kind and r['mode']=='Internal']; return xs[0] if len(xs)==1 else None
    def objects(part):
      root=E.fromstring(z.read(part)); result=[]
      for sp in root.iter('{%s}sp'%P):
        ph=sp.find('{%s}nvSpPr/{%s}nvPr/{%s}ph'%(P,P,P)); name=sp.find('{%s}nvSpPr/{%s}cNvPr'%(P,P))
        if ph is not None: result.append({'name':name.get('name') if name is not None else None,'type':ph.get('type'), 'index':ph.get('idx'), 'text':' '.join(n.text or '' for n in sp.iter('{%s}t'%A))})
      return result
    pages=[]
    for i in (1,2):
      part=f'ppt/slides/slide{i}.xml'; layout=one(part,'slideLayout'); note=one(part,'notesSlide')
      pages.append({'part':part,'layout':layout,'slidePlaceholders':objects(part),'layoutPlaceholders':objects(layout) if layout else [],'note':note,'noteBacklink':one(note,'slide') if note else None,'noteText':' '.join(n.text or '' for n in E.fromstring(z.read(note)).iter('{%s}t'%A)) if note else None})
    return pages
rounds=[inspect(f) for f in sys.argv[1:-1]]
first=rounds[0]
for i,page in enumerate(first):
  assert [(x['type'],x['index'],x['text']) for x in page['slidePlaceholders']]==[('title','0',f'Title {chr(65+i)}'),('body','4294967295',f'Body {chr(65+i)}')]
  assert [(x['type'],x['index']) for x in page['layoutPlaceholders']]==[('title','0'),('body','4294967295')]
  assert page['noteText']==f'Notes {chr(65+i)}' and page['noteBacklink']==page['part']
assert first[0]['layout']!=first[1]['layout']
doc=fitz.open(sys.argv[-1]); pdfText=[p.get_text() for p in doc]
assert len(pdfText)==2 and all(f'Title {chr(65+i)}' in text and f'Body {chr(65+i)}' in text for i,text in enumerate(pdfText))
print(json.dumps({'oracle':'Python zipfile, ElementTree and PyMuPDF','rounds':rounds,'finalPdfText':pdfText}))
`;
const independent = JSON.parse((await run('python3', ['-c', python, ...files, pdf], opts)).stdout);
const hash = bytes => createHash('sha256').update(bytes).digest('hex');
const fileHashes = await Promise.all(files.map(async file => ({ file, sha256: hash(await readFile(file)) })));
const evidence = { schemaVersion: 1, scope: 'Two copied slides mapped to distinct layouts with explicit indexes 0 and 4294967295; LibreOfficeDev headless three saves and final PDF reopen. No interactive edits, repair-dialog, PowerPoint or full visual fidelity claim.', environment: { application, platform: platform(), release: release(), arch: arch() }, directory, fileHashes, pdf: { file: pdf, sha256: hash(await readFile(pdf)) }, independent, logs, pdfLog };
await writeFile('docs/testing/copy-placeholder-index-libreoffice-evidence.json', JSON.stringify(evidence, null, 2) + '\n');
console.log(JSON.stringify({ directory, application, indexes: independent.rounds.map(round => round.map(page => ({ slide: page.slidePlaceholders.map(p => [p.type, p.index]), layout: page.layoutPlaceholders.map(p => [p.type, p.index]), notes: page.noteText }))), finalPdfText: independent.finalPdfText }, null, 2));
