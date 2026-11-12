// Optional headless LibreOffice producer evidence; this is not interactive validation.
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

if (process.argv.length > 3) throw new Error('Usage: node scripts/inspect-picture-save.mjs [soffice-executable]');
const executable = process.argv[2] ?? 'soffice', execute = promisify(execFile), opts = { timeout: 60000, maxBuffer: 8 * 1024 * 1024 };
const application = (await execute(executable, ['--version'], opts)).stdout.trim();
await mkdir('.tmp/picture-save-native', { recursive: true });
const directory = await mkdtemp(path.resolve('.tmp/picture-save-native/run-')), pdfDir = path.join(directory, 'pdf'); await mkdir(pdfDir);
const replacement = await readFile('tests/fixtures/transparent-replacement.png'), doc = createPresentation({ size: { width: 500, height: 250 } });
doc.registerAsset({ id: 'shared', kind: 'image', source: { type: 'bytes', value: png }, mimeType: 'image/png', width: 8, height: 4 });
doc.addSlide({ id: 'slide', elements: [
  { id: 'target', type: 'image', assetId: 'shared', fit: 'crop', crop: { left: 0.125 }, box: { x: 30, y: 30, width: 150, height: 150 } },
  { id: 'sibling', type: 'image', assetId: 'shared', fit: 'crop', crop: { left: 0.125 }, box: { x: 300, y: 30, width: 150, height: 150 } },
] });
const source = (await generatePptx(doc)).bytes, session = await importPptx(source, { projection: 'content', pictureEditSaveProfile: 'replace-v1' });
let edited, report;
try {
  const slide = session.document.slides[0], target = slide.elements[0];
  session.document.registerAsset({ id: 'replacement', kind: 'image', source: { type: 'bytes', value: replacement }, mimeType: 'image/png', width: 4, height: 2 });
  slide.replaceElement(target.id, { ...target, assetId: 'replacement' });
  const plan = await session.prepareSave({ mode: 'edit' }); assert.equal(plan.status, 'ready', JSON.stringify(plan.report));
  const saved = await session.save(plan.plan); edited = saved.bytes; report = saved.report;
} finally { session.dispose(); }
const sourceFile = path.join(directory, 'source.pptx'), editedFile = path.join(directory, 'edited.pptx');
await writeFile(sourceFile, source); await writeFile(editedFile, edited);
const profile = `-env:UserInstallation=${pathToFileURL(path.join(directory, 'profile')).href}`, logs = [];
for (const file of [sourceFile, editedFile]) logs.push(await execute(executable, [profile, '--headless', '--convert-to', 'pdf:impress_pdf_Export', '--outdir', pdfDir, file], opts));
let current = editedFile; const nativeFiles = [];
for (let round = 1; round <= 3; round++) {
  const folder = path.join(directory, `round-${round}`); await mkdir(folder);
  logs.push(await execute(executable, [profile, '--headless', '--convert-to', 'pptx:Impress Office Open XML', '--outdir', folder, current], opts));
  current = path.join(folder, path.basename(current)); nativeFiles.push(current);
  const pdfFolder = path.join(folder, 'pdf'); await mkdir(pdfFolder);
  logs.push(await execute(executable, [profile, '--headless', '--convert-to', 'pdf:impress_pdf_Export', '--outdir', pdfFolder, current], opts));
}
const python = String.raw`
import sys,json,zipfile,xml.etree.ElementTree as ET,fitz,hashlib,os
source,edited,*native=sys.argv[1:]
def pix(file,rect):
  doc=fitz.open(file); assert len(doc)==1
  clip=fitz.Rect(*rect); assert doc[0].rect.contains(clip)
  return hashlib.sha256(doc[0].get_pixmap(matrix=fitz.Matrix(2,2),clip=clip,alpha=False).samples).hexdigest()
def pictures(file):
  with zipfile.ZipFile(file) as z:
    p='{http://schemas.openxmlformats.org/presentationml/2006/main}'
    a='{http://schemas.openxmlformats.org/drawingml/2006/main}'
    r='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
    root=ET.fromstring(z.read('ppt/slides/slide1.xml'))
    rels=ET.fromstring(z.read('ppt/slides/_rels/slide1.xml.rels'))
    targets={n.get('Id'):n.get('Target') for n in rels}
    return [{'name': n.find('.//'+p+'cNvPr').get('name'), 'embed': n.find('.//'+a+'blip').get(r+'embed'), 'target': targets[n.find('.//'+a+'blip').get(r+'embed')], 'crop': dict(n.find('.//'+a+'srcRect').attrib) if n.find('.//'+a+'srcRect') is not None else None} for n in root.iter(p+'pic')]
pdfs=[os.path.join(os.path.dirname(source),'pdf','source.pdf'),os.path.join(os.path.dirname(edited),'pdf','edited.pdf')]+[os.path.join(os.path.dirname(file),'pdf','edited.pdf') for file in native]
regions={'target':(30,30,180,180),'sibling':(300,30,450,180)}
pixels=[{key:pix(pdf,rect) for key,rect in regions.items()} for pdf in pdfs]
assert pixels[0]['target']!=pixels[1]['target'], 'replacement did not change target PDF region'
assert pixels[0]['sibling']==pixels[1]['sibling'], 'replacement changed sibling PDF region'
records=[pictures(file) for file in [source,edited,*native]]
assert all(len(record)==2 for record in records)
assert all([item['name'] for item in record]==['target','sibling'] for record in records)
assert records[0][0]['target']==records[0][1]['target']
assert all(record[0]['target']!=record[1]['target'] for record in records[1:])
assert [int(item['crop']['l']) for item in records[0]]==[12500,12500]
assert [int(item['crop']['l']) for item in records[1]]==[12500,12500]
crop_values=[[int(item['crop']['l']) for item in record] for record in records[2:]]
print(json.dumps({'oracle':'Python zipfile + ElementTree + PyMuPDF','pictureRecords':records,'regionSha256':pixels,'targetChangedInEditedPdf':True,'siblingUnchangedInEditedPdf':True,'targetMatchesEditedAfterNativeRounds':[p['target']==pixels[1]['target'] for p in pixels[2:]],'siblingMatchesEditedAfterNativeRounds':[p['sibling']==pixels[1]['sibling'] for p in pixels[2:]],'nativeCropLeftValues':crop_values,'nativeCropLeftDeltaFromEdited':[[value-12500 for value in row] for row in crop_values]}))
`;
const independent = JSON.parse((await execute('python3', ['-c', python, sourceFile, editedFile, ...nativeFiles], opts)).stdout);
const hash = bytes => createHash('sha256').update(bytes).digest('hex');
const evidence = { schemaVersion: 1, scope: 'One shared two-picture slide, one PPTKit replacement with preserved crop, LibreOffice headless PDF render and three producer resaves. No interactive editing, repair-dialog observation, PowerPoint, arbitrary effects or full image-save conformance.', environment: { platform: platform(), release: release(), arch: arch(), application }, directory, sourceSha256: hash(source), editedSha256: hash(edited), nativeSha256: await Promise.all(nativeFiles.map(async file => hash(await readFile(file)))), report, independent, logs };
await writeFile('docs/testing/picture-save-libreoffice-evidence.json', JSON.stringify(evidence, null, 2) + '\n');
console.log(JSON.stringify({ directory, application, independent }, null, 2));
