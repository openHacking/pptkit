// Controlled producer comparison: placeholder text and ordinary text links on the same pages.
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

const run = promisify(execFile), executable = process.argv[2] ?? 'soffice', execOptions = { timeout: 60000, maxBuffer: 8 * 1024 * 1024 };
const application = (await run(executable, ['--version'], execOptions)).stdout.trim();
const box = { x: 60, y: 40, width: 400, height: 60 }, source = createPresentation();
source.defineSlideLayout({ id: 'layout', name: 'Source', placeholders: [{ key: 'title', kind: 'title', box }] });
for (const [id, other] of [['a', 'b'], ['b', 'a']]) source.addSlide({ id, layoutId: 'layout', notes: `Notes ${id.toUpperCase()}`, elements: [
  { id: `ph-${id}`, name: `placeholder-${id}`, type: 'text', placeholderKey: 'title', box, content: `Placeholder ${id.toUpperCase()}`, action: { type: 'slide', slideId: other } },
  { id: `control-${id}`, name: `control-${id}`, type: 'text', box: { ...box, y: 150 }, content: `Control ${id.toUpperCase()}`, action: { type: 'slide', slideId: other } },
] });
const dest = createPresentation(), assets = { deduplicate: 'content', external: 'reject' };
for (const id of ['a', 'b']) dest.defineSlideLayout({ id: `layout-${id}`, name: id, placeholders: [{ key: `title-${id}`, kind: 'title', box }] });
dest.copySlidesFrom(source, ['a', 'b'], { layoutId: null, layouts: { a: 'layout-a', b: 'layout-b' }, theme: 'destination', geometry: 'preserve', customData: 'preserve', assets, perSlide: { a: { placeholders: { title: 'title-a' } }, b: { placeholders: { title: 'title-b' } } } });
assert.deepEqual(validatePresentation(dest), []); source.removeSlide('a'); source.removeSlide('b');
await mkdir('.tmp/copy-placeholder-links-libreoffice', { recursive: true });
const directory = await mkdtemp(path.resolve('.tmp/copy-placeholder-links-libreoffice/run-'));
let current = path.join(directory, 'input.pptx'); await writeFile(current, (await generatePptx(dest)).bytes);
const files = [current], logs = [], profile = `-env:UserInstallation=${pathToFileURL(path.join(directory, 'profile')).href}`;
for (let round = 1; round <= 3; round++) {
  const outdir = path.join(directory, `round-${round}`); await mkdir(outdir);
  const result = await run(executable, [profile, '--headless', '--convert-to', 'pptx:Impress Office Open XML', '--outdir', outdir, current], execOptions);
  logs.push({ stdout: result.stdout, stderr: result.stderr }); current = path.join(outdir, 'input.pptx'); await readFile(current); files.push(current);
}
const pdfDir = path.join(directory, 'pdf'); await mkdir(pdfDir);
const pdfLog = await run(executable, [profile, '--headless', '--convert-to', 'pdf:impress_pdf_Export', '--outdir', pdfDir, current], execOptions);
const pdf = path.join(pdfDir, 'input.pdf'); await readFile(pdf);
const python = String.raw`
import sys,json,hashlib,zipfile,xml.etree.ElementTree as E,fitz,posixpath
P='http://schemas.openxmlformats.org/presentationml/2006/main'; A='http://schemas.openxmlformats.org/drawingml/2006/main'
R='http://schemas.openxmlformats.org/package/2006/relationships'; O='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
def inspect(file):
  with zipfile.ZipFile(file) as z:
    pages=[]
    for index in (1,2):
      part='ppt/slides/slide%d.xml'%index; root=E.fromstring(z.read(part)); rel='ppt/slides/_rels/slide%d.xml.rels'%index
      relationships={x.get('Id'):x for x in E.fromstring(z.read(rel)) if x.tag=='{%s}Relationship'%R}
      shapes=[]
      for shape in root.iter('{%s}sp'%P):
        nv=shape.find('{%s}nvSpPr/{%s}cNvPr'%(P,P)); placeholder=shape.find('{%s}nvSpPr/{%s}nvPr/{%s}ph'%(P,P,P)); click=nv.find('{%s}hlinkClick'%A) if nv is not None else None
        rid=click.get('{%s}id'%O) if click is not None else None; reln=relationships.get(rid)
        target=posixpath.normpath(posixpath.join('ppt/slides',reln.get('Target'))) if reln is not None and reln.get('TargetMode','Internal')=='Internal' else None
        shapes.append({'name':nv.get('name') if nv is not None else None,'text':''.join(t.text or '' for t in shape.iter('{%s}t'%A)), 'placeholder':placeholder is not None,'clickId':rid,'target':target,'type':reln.get('Type','').split('/')[-1] if reln is not None else None})
      pages.append(shapes)
    return pages
rounds=[inspect(f) for f in sys.argv[1:-1]]
def byText(page,label):
  found=[s for s in page if s['text']==label]
  return found[0] if len(found)==1 else None
def check(pages,index,prefix,target):
  shape=byText(pages[index],prefix+(' A' if index==0 else ' B'))
  return shape is not None and shape['type']=='slide' and shape['target']==target
assert all(check(rounds[0],i,p,t) for i,t in ((0,'ppt/slides/slide2.xml'),(1,'ppt/slides/slide1.xml')) for p in ('Placeholder','Control'))
assert all(byText(rounds[0][i],'Placeholder '+('A' if i==0 else 'B'))['placeholder'] for i in (0,1))
assert all(not byText(rounds[0][i],'Control '+('A' if i==0 else 'B'))['placeholder'] for i in (0,1))
document=fitz.open(sys.argv[-1]); pages=[p.get_text() for p in document]
assert len(pages)==2 and all(label in pages[i] for i in (0,1) for label in ('Placeholder '+('A' if i==0 else 'B'),'Control '+('A' if i==0 else 'B')))
comparisons=[{'placeholderLinks':[check(item,i,'Placeholder',t) for i,t in ((0,'ppt/slides/slide2.xml'),(1,'ppt/slides/slide1.xml'))], 'controlLinks':[check(item,i,'Control',t) for i,t in ((0,'ppt/slides/slide2.xml'),(1,'ppt/slides/slide1.xml'))]} for item in rounds]
print(json.dumps({'oracle':'Python zipfile, ElementTree and PyMuPDF','rounds':rounds,'comparisons':comparisons,'reopenedPdfPageText':pages}))
`;
const independent = JSON.parse((await run('python3', ['-c', python, ...files, pdf], execOptions)).stdout);
const hash = bytes => createHash('sha256').update(bytes).digest('hex');
const fileHashes = await Promise.all(files.map(async file => ({ file, sha256: hash(await readFile(file)) })));
const evidence = { schemaVersion: 1, scope: 'Same-page placeholder and ordinary text link controls in a two-page cross-document batch; LibreOfficeDev headless three saves and final PDF reopen. Not interactive activation, repair-dialog, PowerPoint or complete visual fidelity.', environment: { application, platform: platform(), release: release(), arch: arch() }, directory, fileHashes, pdf: { file: pdf, sha256: hash(await readFile(pdf)) }, independent, logs, pdfLog };
await writeFile('docs/testing/copy-placeholder-links-libreoffice-evidence.json', JSON.stringify(evidence, null, 2) + '\n');
console.log(JSON.stringify({ directory, application, comparisons: independent.comparisons }, null, 2));
