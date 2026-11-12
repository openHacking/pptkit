// Scoped, headless producer observation; this does not establish interactive behavior.
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
import { importPptx } from '../packages/pptx-parser/dist/index.js';
import { png } from '../packages/pptx-parser/test/image-fixture.mjs';

const run = promisify(execFile), executable = process.argv[2] ?? 'soffice';
const execOptions = { timeout: 60000, maxBuffer: 8 * 1024 * 1024 };
const application = (await run(executable, ['--version'], execOptions)).stdout.trim();
const source = createPresentation(), dest = createPresentation(), box = { x: 60, y: 45, width: 300, height: 60 };
source.registerAsset({ id: 'image', kind: 'image', mimeType: 'image/png', source: { type: 'bytes', value: png }, width: 8, height: 4 });
for (const [id, other] of [['a', 'b'], ['b', 'a']]) source.addSlide({ id, notes: [{ runs: [{ text: `Notes ${id.toUpperCase()}`, action: { type: 'slide', slideId: other } }] }], elements: [
  { id: `${id}-text`, type: 'text', content: `Page ${id.toUpperCase()}`, box, action: { type: 'slide', slideId: other } },
  { id: `${id}-image`, type: 'image', assetId: 'image', box: { x: 80, y: 150, width: 240, height: 120 } },
] });
dest.addSlide({ id: 'control', notes: 'Control notes', elements: [{ id: 'control-text', type: 'text', content: 'Control page', box }] });
const copied = dest.copySlidesFrom(source, ['b', 'a'], { layoutId: null, theme: 'destination', geometry: 'preserve', customData: 'preserve', assets: { deduplicate: 'content', external: 'reject' } });
assert.equal(copied.slides.length, 2); source.removeSlide('a'); source.removeSlide('b');
await mkdir('.tmp/cross-batch-libreoffice', { recursive: true });
const directory = await mkdtemp(path.resolve('.tmp/cross-batch-libreoffice/run-'));
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
    notes=sorted((n for n in names if re.fullmatch(r'ppt/notesSlides/notesSlide\d+\.xml',n)),key=lambda n:int(re.search(r'\d+',n).group()))
    def text(root): return ''.join(n.text or '' for n in root.iter('{%s}t'%A))
    def links(part):
      root=E.fromstring(z.read(part)); rel=posixpath.join(posixpath.dirname(part),'_rels',posixpath.basename(part)+'.rels')
      relationships={x.get('Id'):x for x in E.fromstring(z.read(rel)) if x.tag=='{%s}Relationship'%R}; result=[]
      for click in root.iter('{%s}hlinkClick'%A):
        rid=click.get('{%s}id'%O); item=relationships.get(rid)
        raw=item.get('Target') if item is not None else None; mode=item.get('TargetMode','Internal') if item is not None else None
        result.append({'id':rid,'rawTarget':raw,'targetMode':mode,'resolvedPart':posixpath.normpath(posixpath.join(posixpath.dirname(part),raw)) if raw is not None and mode=='Internal' else None,'type':item.get('Type','').split('/')[-1] if item is not None else None})
      return result
    return {'slides':[{'part':n,'text':text(E.fromstring(z.read(n))),'clicks':links(n)} for n in slides], 'notes':[{'part':n,'text':text(E.fromstring(z.read(n))),'clicks':links(n)} for n in notes],
      'media':[{'part':n,'sha256':hashlib.sha256(z.read(n)).hexdigest()} for n in names if n.startswith('ppt/media/') and not n.endswith('/')]}
items=[inspect(f) for f in sys.argv[1:-1]]
assert [v['text'] for v in items[0]['slides']]==['Control page','Page B','Page A']
assert [v['text'] for v in items[0]['notes']]==['Control notes','Notes B','Notes A']
assert items[0]['slides'][1]['clicks'][0]['resolvedPart']=='ppt/slides/slide3.xml' and items[0]['slides'][2]['clicks'][0]['resolvedPart']=='ppt/slides/slide2.xml'
assert items[0]['notes'][1]['clicks'][0]['resolvedPart']=='ppt/slides/slide3.xml' and items[0]['notes'][2]['clicks'][0]['resolvedPart']=='ppt/slides/slide2.xml'
assert len(items[0]['media'])==1
document=fitz.open(sys.argv[-1]); pages=[p.get_text() for p in document]
assert len(pages)==3 and all(label in text for label,text in zip(['Control page','Page B','Page A'],pages))
def expectedLinks(entries):
  if len(entries)!=3: return False
  return all(len(entries[i]['clicks'])==1 and entries[i]['clicks'][0]['type']=='slide' and entries[i]['clicks'][0]['targetMode']=='Internal' and entries[i]['clicks'][0]['resolvedPart']==target for i,target in ((1,'ppt/slides/slide3.xml'),(2,'ppt/slides/slide2.xml')))
baseline=items[0]['media'][0]['sha256']
comparisons=[{'pageTextRetained':[s['text'] for s in item['slides']]==['Control page','Page B','Page A'], 'notesTextRetained':[n['text'] for n in item['notes']]==['Control notes','Notes B','Notes A'], 'slideLinksRetained':expectedLinks(item['slides']), 'notesLinksRetained':expectedLinks(item['notes']), 'oneIdenticalImage':len(item['media'])==1 and item['media'][0]['sha256']==baseline} for item in items]
print(json.dumps({'oracle':'Python zipfile, ElementTree and PyMuPDF','rounds':items,'comparisons':comparisons,'reopenedPdfPageText':pages}))
`;
const independent = JSON.parse((await run('python3', ['-c', python, ...files, pdf], execOptions)).stdout);
const hash = bytes => createHash('sha256').update(bytes).digest('hex');
const fileHashes = await Promise.all(files.map(async file => ({ file, sha256: hash(await readFile(file)) })));
const parser = [];
for (const [index, file] of files.entries()) {
  const bytes = await readFile(file), session = await importPptx(bytes, { projection: 'content' });
  try {
    const slides = session.document.slides;
    const noteActions = slides.map(slide => typeof slide.notes === 'string' || slide.notes === undefined ? [] : slide.notes.flatMap(paragraph => paragraph.runs.flatMap(run => run.action ? [run.action] : [])));
    const diagnostics = session.report.diagnostics.filter(item => item.code === 'notes-click-action-unmapped').map(item => ({ slideId: item.slideId, message: item.message }));
    const plan = await session.prepareSave({ mode: 'no-op' }); assert.equal(plan.status, 'ready');
    assert.deepEqual((await session.save(plan.plan)).bytes, new Uint8Array(bytes));
    if (index === 0) { assert.equal(noteActions[1][0]?.slideId, slides[2].id); assert.equal(noteActions[2][0]?.slideId, slides[1].id); assert.equal(diagnostics.length, 0); }
    else { assert.deepEqual(noteActions, [[], [], []]); assert.equal(diagnostics.length, 2); }
    parser.push({ slideIds: slides.map(slide => slide.id), noteActions, diagnostics, noOpExact: true });
  } finally { session.dispose(); }
}
const evidence = { schemaVersion: 1, scope: 'One control and two mutually linked copied pages with linked notes and a shared picture; LibreOfficeDev headless three saves and final PDF reopen. This is not interactive editability, link activation, repair-dialog, PowerPoint or full visual fidelity.', environment: { application, platform: platform(), release: release(), arch: arch() }, directory, fileHashes, pdf: { file: pdf, sha256: hash(await readFile(pdf)) }, independent, parser, logs, pdfLog };
await writeFile('docs/testing/cross-batch-libreoffice-evidence.json', JSON.stringify(evidence, null, 2) + '\n');
console.log(JSON.stringify({ directory, application, rounds: independent.rounds.map(round => ({ slideClicks: round.slides.map(slide => slide.clicks), notesClicks: round.notes.map(note => note.clicks), media: round.media })), pdfPages: independent.reopenedPdfPageText.length }, null, 2));
