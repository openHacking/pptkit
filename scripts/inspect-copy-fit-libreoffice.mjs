// Optional headless producer observation; no interactive editability claim.
import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { createHash } from 'node:crypto';
import { mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises';
import { arch, platform, release } from 'node:os';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { createPresentation } from '../packages/core/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';

const run = promisify(execFile), executable = process.argv[2] ?? 'soffice';
const opts = { timeout: 60000, maxBuffer: 8 * 1024 * 1024 };
const application = (await run(executable, ['--version'], opts)).stdout.trim();
const source = createPresentation({ size: { width: 400, height: 200 } });
source.addSlide({ id: 'source', notes: 'Fit notes', elements: [
  { id: 'heading', type: 'text', content: 'Geometry fit', box: { x: 20, y: 10, width: 160, height: 20 } },
  { id: 'shape', type: 'shape', shape: 'rect', box: { x: 20, y: 40, width: 60, height: 30 }, style: { fill: { type: 'solid', color: 'FF0000' } } },
  { id: 'edge', type: 'connector', start: { elementId: 'shape', anchor: 'right' }, end: { x: 200, y: 55 }, route: [{ x: 120, y: 75 }] },
  { id: 'group', type: 'group', box: { x: 220, y: 35, width: 100, height: 40 }, coordinateSize: { width: 100, height: 40 }, children: [
    { id: 'child', type: 'shape', shape: 'rect', box: { x: 10, y: 5, width: 40, height: 20 }, style: { fill: { type: 'solid', color: '0000FF' } } },
  ] },
  { id: 'table', type: 'table', box: { x: 20, y: 100, width: 160, height: 40 }, columns: [60, 100], rows: [{ height: 40, cells: [{ content: 'Left' }, { content: 'Right' }] }] },
] });
const destination = createPresentation({ size: { width: 200, height: 200 } });
destination.copySlideFrom(source, 'source', { layoutId: null, theme: 'destination', geometry: 'fit-explicit-geometry', customData: 'preserve', assets: { deduplicate: 'content', external: 'reject' } });
const controlBytes = (await generatePptx(source)).bytes;
source.removeSlide('source');
await mkdir('.tmp/copy-fit-libreoffice', { recursive: true });
const directory = await mkdtemp(path.resolve('.tmp/copy-fit-libreoffice/run-'));
const control = path.join(directory, 'source-control.pptx'); await writeFile(control, controlBytes);
let current = path.join(directory, 'input.pptx'); await writeFile(current, (await generatePptx(destination)).bytes);
const files = [current], logs = [], profile = `-env:UserInstallation=${pathToFileURL(path.join(directory, 'profile')).href}`;
for (let round = 1; round <= 3; round++) {
  const outdir = path.join(directory, `round-${round}`); await mkdir(outdir);
  const result = await run(executable, [profile, '--headless', '--convert-to', 'pptx:Impress Office Open XML', '--outdir', outdir, current], opts);
  logs.push({ stdout: result.stdout, stderr: result.stderr }); current = path.join(outdir, 'input.pptx'); await readFile(current); files.push(current);
}
const inputPdfDir = path.join(directory, 'input-pdf'); await mkdir(inputPdfDir);
const inputPdfLog = await run(executable, [profile, '--headless', '--convert-to', 'pdf:impress_pdf_Export', '--outdir', inputPdfDir, files[0]], opts);
const inputPdf = path.join(inputPdfDir, 'input.pdf'); await readFile(inputPdf);
const controlPdfDir = path.join(directory, 'control-pdf'); await mkdir(controlPdfDir);
const controlPdfLog = await run(executable, [profile, '--headless', '--convert-to', 'pdf:impress_pdf_Export', '--outdir', controlPdfDir, control], opts);
const controlPdf = path.join(controlPdfDir, 'source-control.pdf'); await readFile(controlPdf);
const outdir = path.join(directory, 'pdf'); await mkdir(outdir);
const pdfLog = await run(executable, [profile, '--headless', '--convert-to', 'pdf:impress_pdf_Export', '--outdir', outdir, current], opts);
const pdf = path.join(outdir, 'input.pdf'); await readFile(pdf);
const python = String.raw`
import sys,json,zipfile,xml.etree.ElementTree as E,fitz
P='http://schemas.openxmlformats.org/presentationml/2006/main'; A='http://schemas.openxmlformats.org/drawingml/2006/main'
def inspect(file):
  with zipfile.ZipFile(file) as z:
    slide=E.fromstring(z.read('ppt/slides/slide1.xml')); tree=slide.find('.//{%s}spTree'%P)
    def node(e):
      name=e.find('.//{%s}cNvPr'%P); kind=e.tag.split('}')[-1]
      trans=e.find('{%s}spPr/{%s}xfrm'%(P,A)) if kind in ('sp','pic','cxnSp') else e.find('{%s}grpSpPr/{%s}xfrm'%(P,A)) if kind=='grpSp' else e.find('{%s}xfrm'%P)
      geom={}
      if trans is not None:
        for k in ('off','ext','chOff','chExt'):
          a=trans.find('{%s}%s'%(A,k))
          if a is not None: geom[k]=dict(a.attrib)
      return {'kind':kind,'name':name.get('name') if name is not None else None,'geometry':geom,
        'text':' '.join(x.text or '' for x in e.iter('{%s}t'%A)),
        'grid':[int(x.get('w')) for x in e.iter('{%s}gridCol'%A)],
        'rows':[int(x.get('h')) for x in e.iter('{%s}tr'%A)],
        'children':[node(x) for x in e if x.tag.split('}')[-1] in ('sp','pic','cxnSp','grpSp','graphicFrame')] if kind=='grpSp' else []}
    objects=[node(e) for e in tree if e.tag.split('}')[-1] in ('sp','pic','cxnSp','grpSp','graphicFrame')]
    text=' '.join(x.text or '' for x in slide.iter('{%s}t'%A)); notes=' '.join(x.text or '' for x in E.fromstring(z.read('ppt/notesSlides/notesSlide1.xml')).iter('{%s}t'%A))
    return {'objects':objects,'text':text,'notes':notes}
items=[inspect(f) for f in sys.argv[1:-3]]
objects=items[0]['objects']; assert len(objects)==5
assert objects[0]['geometry']['off']=={'x':'127000','y':'698500'}
assert objects[1]['geometry']['off']=={'x':'127000','y':'889000'}
assert objects[1]['geometry']['ext']=={'cx':'381000','cy':'190500'}
assert objects[3]['geometry']['off']=={'x':'1397000','y':'857250'}
assert objects[3]['geometry']['ext']=={'cx':'635000','cy':'254000'}
assert objects[4]['grid']==[381000,635000] and objects[4]['rows']==[254000]
for x in items: assert all(s in x['text'] for s in ('Geometry fit','Left','Right')) and 'Fit notes' in x['notes']
def pdf(file):
  doc=fitz.open(file); assert len(doc)==1
  page=doc[0]; text=page.get_text(); compact=''.join(text.split()); assert all(s in compact for s in ('Geometryfit','Left','Right'))
  page.get_pixmap(matrix=fitz.Matrix(3,3), alpha=False).save(file+'.png')
  return {'pageSize':[page.rect.width,page.rect.height], 'text':text, 'continuousWordsPresent':[s in text for s in ('Geometry fit','Left','Right')], 'drawingCount':len(page.get_drawings())}
control=pdf(sys.argv[-3]); initial=pdf(sys.argv[-2]); final=pdf(sys.argv[-1])
a=fitz.Pixmap(sys.argv[-2]+'.png'); b=fitz.Pixmap(sys.argv[-1]+'.png')
assert (a.width,a.height,a.n)==(b.width,b.height,b.n)
aa,bb=a.samples,b.samples
changed=sum(aa[i:i+a.n]!=bb[i:i+b.n] for i in range(0,len(aa),a.n))
print(json.dumps({'oracle':'Python zipfile, ElementTree and PyMuPDF','rounds':items,'controlPdf':control,'inputPdf':initial,'finalPdf':final,'rasterComparison':{'width':a.width,'height':a.height,'changedPixels':changed,'totalPixels':a.width*a.height}}))
`;
const independent = JSON.parse((await run('python3', ['-c', python, ...files, controlPdf, inputPdf, pdf], opts)).stdout);
const hash = bytes => createHash('sha256').update(bytes).digest('hex');
const fileHashes = await Promise.all(files.map(async file => ({ file, sha256: hash(await readFile(file)) })));
const evidence = { schemaVersion: 1, scope: 'One explicit-geometry fit from 400x200 to 200x200 points, with text, rectangle, static connector, group and table; LibreOfficeDev headless three saves and source, input and final PDF conversions. No interactive editing, PowerPoint, or complete visual fidelity claim.', environment: { application, platform: platform(), release: release(), arch: arch() }, directory, fileHashes, control: { file: control, sha256: hash(await readFile(control)), pdf: controlPdf, pdfSha256: hash(await readFile(controlPdf)), raster: `${controlPdf}.png`, rasterSha256: hash(await readFile(`${controlPdf}.png`)) }, inputPdf: { file: inputPdf, sha256: hash(await readFile(inputPdf)), raster: `${inputPdf}.png`, rasterSha256: hash(await readFile(`${inputPdf}.png`)) }, pdf: { file: pdf, sha256: hash(await readFile(pdf)), raster: `${pdf}.png`, rasterSha256: hash(await readFile(`${pdf}.png`)) }, independent, logs, controlPdfLog, inputPdfLog, pdfLog };
await writeFile('docs/testing/copy-fit-libreoffice-evidence.json', JSON.stringify(evidence, null, 2) + '\n');
console.log(JSON.stringify({ directory, application, controlPdf: independent.controlPdf, inputPdf: independent.inputPdf, finalPdf: independent.finalPdf }, null, 2));
