// Optional headless application observation; does not drive PowerPoint.
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { createHash } from 'node:crypto';
import { mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { createPresentation } from '../packages/core/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';

const execute = promisify(execFile), options = { timeout: 60000, maxBuffer: 8 * 1024 * 1024 }, executable = process.argv[2] ?? 'soffice';
const application = (await execute(executable, ['--version'], options)).stdout.trim();
await mkdir('.tmp/first-slide-number-libreoffice', { recursive: true });
const directory = await mkdtemp(path.resolve('.tmp/first-slide-number-libreoffice/run-'));
const profile = `-env:UserInstallation=${pathToFileURL(path.join(directory, 'profile')).href}`;
const deck = createPresentation();
for (let i = 1; i <= 2; i++) deck.addSlide({ elements: [{ type: 'text', box: { x: 30, y: 30, width: 300, height: 60 }, content: [{ runs: [{ text: `Page ${i}: ` }, { text: 'CACHE', field: { id: `11111111-2222-3333-4444-${String(i).padStart(12, '0')}`, kind: 'slide-number' } }] }] }] });
const source = path.join(directory, 'source.pptx'); await writeFile(source, (await generatePptx(deck, { firstSlideNumber: 17 })).bytes);
const round = path.join(directory, 'round'); await mkdir(round);
await execute(executable, [profile, '--headless', '--convert-to', 'pptx:Impress Office Open XML', '--outdir', round, source], options);
const files = [source, path.join(round, 'source.pptx')];
for (const [index, file] of files.entries()) { const folder = path.join(directory, `pdf-${index}`); await mkdir(folder); await execute(executable, [profile, '--headless', '--convert-to', 'pdf:impress_pdf_Export', '--outdir', folder, file], options); }
const python = `import sys,json,zipfile,xml.etree.ElementTree as E,fitz,os\na='{http://schemas.openxmlformats.org/drawingml/2006/main}'\nout=[]\nfor i,p in enumerate(sys.argv[2:]):\n with zipfile.ZipFile(p) as z:\n  pres=E.fromstring(z.read('ppt/presentation.xml'))\n  fields=[]\n  for n in (1,2):\n   root=E.fromstring(z.read('ppt/slides/slide%d.xml'%n))\n   fields.append([{'id':f.get('id'),'type':f.get('type'),'cache':''.join(t.text or '' for t in f.iter(a+'t'))} for f in root.iter(a+'fld')])\n doc=fitz.open(os.path.join(sys.argv[1],'pdf-'+str(i),'source.pdf'))\n out.append({'firstSlideNum':pres.get('firstSlideNum'),'fields':fields,'pdfText':[p.get_text().strip() for p in doc]})\nprint(json.dumps(out))`;
const records = JSON.parse((await execute('python3', ['-c', python, directory, ...files], options)).stdout);
if (records[0].firstSlideNum !== '17' || records[0].fields.some(fields => fields.length !== 1 || fields[0].cache !== 'CACHE' || fields[0].type !== 'slidenum')) throw Error('Source does not match authored specimen');
const hashes = await Promise.all(files.map(async file => ({ file, sha256: createHash('sha256').update(await readFile(file)).digest('hex') })));
const evidence = { schemaVersion: 1, application, directory, files: hashes, records, scope: 'Two visible slides, firstSlideNum 17 and stale CACHE fields; one LibreOffice headless conversion and PDF render per file. No PowerPoint or interactive claim.' };
await writeFile('docs/testing/first-slide-number-libreoffice-evidence.json', JSON.stringify(evidence, null, 2) + '\n');
console.log(JSON.stringify(evidence));
