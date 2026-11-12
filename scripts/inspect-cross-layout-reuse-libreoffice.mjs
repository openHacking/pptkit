// Headless producer observation for a cross-document master/layout/page chain.
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
import { png } from '../packages/pptx-parser/test/image-fixture.mjs';

const run = promisify(execFile), executable = process.argv[2] ?? 'soffice', execOptions = { timeout: 60000, maxBuffer: 8 * 1024 * 1024 };
const application = (await run(executable, ['--version'], execOptions)).stdout.trim();
const box = { x: 60, y: 50, width: 400, height: 60 }, source = createPresentation({ theme: { colors: { accent2: '245678' } }, textStylePresets: { old: { run: { fontSize: 18 } } } });
source.registerAsset({ id: 'image', kind: 'image', mimeType: 'image/png', source: { type: 'bytes', value: png }, width: 8, height: 4 });
source.defineSlideMaster({ id: 'master', name: 'Source Master', elements: [
  { id: 'image', type: 'image', assetId: 'image', box: { x: 80, y: 150, width: 240, height: 120 } },
  { id: 'accent', type: 'shape', shape: 'rect', box: { x: 500, y: 150, width: 120, height: 120 }, style: { fill: { type: 'solid', color: { theme: 'accent2' } } } },
] });
source.defineSlideLayout({ id: 'layout', name: 'Source Layout', masterId: 'master', placeholders: [{ key: 'title', kind: 'title', box }] });
for (const [id, other] of [['a', 'b'], ['b', 'a']]) source.addSlide({ id, layoutId: 'layout', notes: `Notes ${id.toUpperCase()}`, elements: [{ id: `text-${id}`, type: 'text', placeholderKey: 'title', textStylePreset: 'old', box, content: `Page ${id.toUpperCase()}`, action: { type: 'slide', slideId: other } }] });
const dest = createPresentation({ theme: { colors: { accent2: 'FF00FF' } }, textStylePresets: { first: { run: { fontSize: 20 } }, second: { run: { fontSize: 24 } } } });
const assets = { deduplicate: 'content', external: 'reject' };
dest.copySlideMasterFrom(source, 'master', { id: 'copied-master', theme: 'source-effective', assets });
const first = dest.copySlideLayoutFrom(source, 'layout', { id: 'layout-a', masterId: 'copied-master', theme: 'destination', assets });
const second = dest.copySlideLayoutFrom(source, 'layout', { id: 'layout-b', masterId: 'copied-master', theme: 'destination', assets, reuseEquivalent: true });
assert.equal(second.layout, first.layout); assert.equal(dest.layouts.length, 1);
dest.copySlidesFrom(source, ['a', 'b'], { layoutId: first.layout.id, theme: 'destination', geometry: 'preserve', customData: 'preserve', assets, perSlide: { a: { textPresets: { old: 'first' } }, b: { textPresets: { old: 'second' } } } });
assert.deepEqual(validatePresentation(dest), []); source.removeSlide('a'); source.removeSlide('b');
await mkdir('.tmp/cross-layout-reuse-libreoffice', { recursive: true });
const directory = await mkdtemp(path.resolve('.tmp/cross-layout-reuse-libreoffice/run-'));
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
import sys,json,hashlib,zipfile,xml.etree.ElementTree as E,fitz,re,posixpath
P='http://schemas.openxmlformats.org/presentationml/2006/main'; A='http://schemas.openxmlformats.org/drawingml/2006/main'
R='http://schemas.openxmlformats.org/package/2006/relationships'; O='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
def inspect(file):
  with zipfile.ZipFile(file) as z:
    names=z.namelist(); slides=sorted((n for n in names if re.fullmatch(r'ppt/slides/slide\d+\.xml',n)),key=lambda n:int(re.search(r'\d+',n).group()))
    def rels(part):
      rel=posixpath.join(posixpath.dirname(part),'_rels',posixpath.basename(part)+'.rels')
      if rel not in names: return []
      return [{'id':x.get('Id'),'type':x.get('Type','').split('/')[-1], 'target':posixpath.normpath(posixpath.join(posixpath.dirname(part),x.get('Target'))) if x.get('TargetMode','Internal')=='Internal' else None} for x in E.fromstring(z.read(rel)) if x.tag=='{%s}Relationship'%R]
    def one(items,kind):
      found=[x['target'] for x in items if x['type']==kind and x['target'] is not None]
      return found[0] if len(found)==1 else None
    result=[]
    for part in slides:
      root=E.fromstring(z.read(part)); slideRels=rels(part); layout=one(slideRels,'slideLayout'); master=one(rels(layout),'slideMaster') if layout else None; theme=one(rels(master),'theme') if master else None
      themeRoot=E.fromstring(z.read(theme)) if theme else None; accent=themeRoot.find('.//{%s}clrScheme/{%s}accent2/*'%(A,A)) if themeRoot is not None else None
      clicks=[]
      for click in root.iter('{%s}hlinkClick'%A):
        rid=click.get('{%s}id'%O); found=[r for r in slideRels if r['id']==rid]; clicks.append(found[0] if len(found)==1 else None)
      masterRoot=E.fromstring(z.read(master)) if master else None; layoutRoot=E.fromstring(z.read(layout)) if layout else None
      result.append({'part':part,'text':''.join(n.text or '' for n in root.iter('{%s}t'%A)), 'runSizes':[n.get('sz') for n in root.iter('{%s}rPr'%A)], 'layout':layout,'master':master,'theme':theme,'accent2':accent.get('val') if accent is not None else None,'masterPictures':len(list(masterRoot.iter('{%s}pic'%P))) if masterRoot is not None else None,'layoutPictures':len(list(layoutRoot.iter('{%s}pic'%P))) if layoutRoot is not None else None,'clicks':clicks})
    notes=sorted((n for n in names if re.fullmatch(r'ppt/notesSlides/notesSlide\d+\.xml',n)),key=lambda n:int(re.search(r'\d+',n).group()))
    return {'pages':result,'notes':[{'text':''.join(n.text or '' for n in E.fromstring(z.read(part)).iter('{%s}t'%A)), 'slide':one(rels(part),'slide')} for part in notes], 'media':[{'part':n,'sha256':hashlib.sha256(z.read(n)).hexdigest()} for n in names if n.startswith('ppt/media/') and not n.endswith('/')]}
rounds=[inspect(f) for f in sys.argv[1:-1]]; first=rounds[0]
assert [p['text'] for p in first['pages']]==['Page A','Page B']
assert [n['text'] for n in first['notes']]==['Notes A','Notes B'] and [n['slide'] for n in first['notes']]==['ppt/slides/slide1.xml','ppt/slides/slide2.xml']
assert first['pages'][0]['layout']==first['pages'][1]['layout'] and first['pages'][0]['master']==first['pages'][1]['master']
assert all(p['accent2']=='245678' and p['masterPictures']==1 for p in first['pages']) and len(first['media'])==1
assert first['pages'][0]['clicks'][0]['target']=='ppt/slides/slide2.xml' and first['pages'][1]['clicks'][0]['target']=='ppt/slides/slide1.xml'
document=fitz.open(sys.argv[-1]); pages=[p.get_text() for p in document]
assert len(pages)==2 and 'Page A' in pages[0] and 'Page B' in pages[1]
comparisons=[{'textsRetained':[p['text'] for p in item['pages']]==['Page A','Page B'], 'notesAndBacklinksRetained':[n['text'] for n in item['notes']]==['Notes A','Notes B'] and [n['slide'] for n in item['notes']]==['ppt/slides/slide1.xml','ppt/slides/slide2.xml'], 'runSizesRetained':[p['runSizes'] for p in item['pages']]==[['2000'],['2400']], 'sharedLayout':len(item['pages'])==2 and item['pages'][0]['layout'] is not None and item['pages'][0]['layout']==item['pages'][1]['layout'], 'sharedMaster':len(item['pages'])==2 and item['pages'][0]['master'] is not None and item['pages'][0]['master']==item['pages'][1]['master'], 'accent2Retained':len(item['pages'])==2 and all(p['accent2']=='245678' for p in item['pages']), 'masterPictureRetained':len(item['pages'])==2 and all(p['masterPictures']==1 for p in item['pages']), 'layoutPictureCounts':[p['layoutPictures'] for p in item['pages']], 'mutualSlideLinksRetained':len(item['pages'])==2 and all(len(item['pages'][i]['clicks'])==1 and item['pages'][i]['clicks'][0] is not None and item['pages'][i]['clicks'][0]['type']=='slide' and item['pages'][i]['clicks'][0]['target']==target for i,target in ((0,'ppt/slides/slide2.xml'),(1,'ppt/slides/slide1.xml'))), 'imageBytesRetained':len(item['media'])==1 and item['media'][0]['sha256']==first['media'][0]['sha256']} for item in rounds]
print(json.dumps({'oracle':'Python zipfile, ElementTree and PyMuPDF','rounds':rounds,'comparisons':comparisons,'reopenedPdfPageText':pages}))
`;
const independent = JSON.parse((await run('python3', ['-c', python, ...files, pdf], execOptions)).stdout);
const hash = bytes => createHash('sha256').update(bytes).digest('hex');
const fileHashes = await Promise.all(files.map(async file => ({ file, sha256: hash(await readFile(file)) })));
const evidence = { schemaVersion: 1, scope: 'Two cross-document pages using one strictly reused layout and one copied effective-theme master; LibreOfficeDev headless three saves and final PDF reopen. No interactive edit, link activation, repair-dialog, PowerPoint or full visual-fidelity claim.', environment: { application, platform: platform(), release: release(), arch: arch() }, directory, fileHashes, pdf: { file: pdf, sha256: hash(await readFile(pdf)) }, independent, logs, pdfLog };
await writeFile('docs/testing/cross-layout-reuse-libreoffice-evidence.json', JSON.stringify(evidence, null, 2) + '\n');
console.log(JSON.stringify({ directory, application, comparisons: independent.comparisons }, null, 2));
