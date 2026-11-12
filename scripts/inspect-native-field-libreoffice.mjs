// Optional producer experiment; does not claim interactive or PowerPoint validation.
import { execFile } from 'node:child_process';
import assert from 'node:assert/strict';
import { promisify } from 'node:util';
import { createHash } from 'node:crypto';
import { mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { createPresentation } from '../packages/core/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';

const execute = promisify(execFile), opts = { timeout: 60000, maxBuffer: 8 * 1024 * 1024 }, executable = process.argv[2] ?? 'soffice';
const application = (await execute(executable, ['--version'], opts)).stdout.trim();
await mkdir('.tmp/native-field-libreoffice', { recursive: true });
const directory = await mkdtemp(path.resolve('.tmp/native-field-libreoffice/run-')), profile = `-env:UserInstallation=${pathToFileURL(path.join(directory, 'profile')).href}`;
const deck = createPresentation({ size: { width: 500, height: 250 } }), field = { id: '38d9dc4a-2e99-e34a-a7a3-8e01bbc23a4a', kind: 'slide-number' };
deck.addSlide({ id: 'a', elements: [{ type: 'text', id: 't', box: { x: 30, y: 30, width: 300, height: 60 }, content: [{ runs: [{ text: 'Page ' }, { text: '42', field }] }] }] });
const source = path.join(directory, 'source.pptx'); await writeFile(source, (await generatePptx(deck)).bytes);
let current = source; const files = [source], logs = [];
for (let round = 1; round <= 3; round++) {
  const folder = path.join(directory, `round-${round}`); await mkdir(folder);
  logs.push(await execute(executable, [profile, '--headless', '--convert-to', 'pptx:Impress Office Open XML', '--outdir', folder, current], opts));
  current = path.join(folder, 'source.pptx'); files.push(current);
}
for (const [index, file] of files.entries()) {
  const folder = path.join(directory, `pdf-${index}`); await mkdir(folder);
  logs.push(await execute(executable, [profile, '--headless', '--convert-to', 'pdf:impress_pdf_Export', '--outdir', folder, file], opts));
}
const python = `import sys,json,zipfile,xml.etree.ElementTree as E,fitz,os
a='{http://schemas.openxmlformats.org/drawingml/2006/main}'
out=[]
for i,p in enumerate(sys.argv[2:]):
 with zipfile.ZipFile(p) as z:
  root=E.fromstring(z.read('ppt/slides/slide1.xml'))
  fields=[{'id':f.get('id'),'type':f.get('type'),'cache':''.join(t.text or '' for t in f.iter(a+'t'))} for f in root.iter(a+'fld')]
 doc=fitz.open(os.path.join(sys.argv[1],'pdf-'+str(i),'source.pdf'))
 out.append({'fields':fields,'pdfText':doc[0].get_text().strip()})
print(json.dumps(out))`;
const records = JSON.parse((await execute('python3', ['-c', python, directory, ...files], opts)).stdout);
assert.equal(records[0].fields[0]?.cache, '42');
for (const record of records) { assert.equal(record.fields.length, 1); assert.equal(record.fields[0].type, 'slidenum'); assert.equal(record.pdfText, 'Page 1'); }
for (const record of records.slice(1)) assert.equal(record.fields[0].cache, '1');
const sha = bytes => createHash('sha256').update(bytes).digest('hex');
const hashes = await Promise.all(files.map(async file => ({ file, sha256: sha(await readFile(file)) })));
const evidence = { schemaVersion: 1, application, directory, files: hashes, records, logs, scope: 'One slide, intentionally stale caller cache 42. Headless LibreOffice conversions only; no interactive edit, repair-dialog observation or PowerPoint claim.' };
await writeFile('docs/testing/native-field-libreoffice-evidence.json', JSON.stringify(evidence, null, 2) + '\n');
console.log(JSON.stringify({ application, directory, records }));
