// Optional native producer experiment; creates a fresh directory and never overwrites source fixtures.
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
const box = (x, y, width, height) => ({ x, y, width, height });
const shape = (id, color, x, y) => ({ id, name: id, type: 'shape', shape: 'rect', box: box(x, y, 60, 60), style: { fill: { type: 'solid', color } }, accessibility: { description: `Description ${id}` } });
const deck = createPresentation({ size: { width: 400, height: 200 } });
const slide = deck.addSlide({ id: 'page', notes: 'Reorder control notes', elements: [
  shape('A', 'D00000', 20, 20), shape('B', '0040D0', 50, 40),
  { id: 'outer', name: 'Outer', type: 'group', box: box(160, 20, 200, 120), coordinateSize: { width: 200, height: 120 }, children: [
    shape('C', '008000', 0, 0), shape('D', 'F0C000', 30, 20),
    { id: 'inner', name: 'Inner', type: 'group', box: box(100, 20, 90, 80), coordinateSize: { width: 90, height: 80 }, children: [shape('E', 'A000A0', 0, 0), shape('F', '0080A0', 20, 10)] },
  ] },
  { id: 'caption', name: 'Caption', type: 'text', box: box(15, 160, 300, 30), content: 'Reorder control text' },
] });
const control = (await generatePptx(deck)).bytes;
slide.reorderElements(['B', 'A', 'outer', 'caption']);
slide.reorderGroupElements('outer', ['D', 'C', 'inner']);
slide.reorderGroupElements('inner', ['F', 'E']);
const edited = (await generatePptx(deck)).bytes;
await mkdir('.tmp/reorder-libreoffice', { recursive: true });
const directory = await mkdtemp(path.resolve('.tmp/reorder-libreoffice/run-'));
const input = path.join(directory, 'input'); await mkdir(input);
const initial = [path.join(input, 'control.pptx'), path.join(input, 'reordered.pptx')];
await writeFile(initial[0], control); await writeFile(initial[1], edited);
const profile = `-env:UserInstallation=${pathToFileURL(path.join(directory, 'profile')).href}`;
const sets = [initial], logs = [];
for (let round = 1; round <= 3; round++) {
  const output = path.join(directory, `round-${round}`); await mkdir(output);
  const result = await run(executable, [profile, '--headless', '--convert-to', 'pptx:Impress Office Open XML', '--outdir', output, ...sets.at(-1)], opts);
  logs.push({ phase: `round-${round}`, stdout: result.stdout, stderr: result.stderr });
  const files = ['control.pptx', 'reordered.pptx'].map(name => path.join(output, name));
  await Promise.all(files.map(file => readFile(file))); sets.push(files);
}
const pdfSets = [];
for (let round = 0; round <= 3; round++) {
  const output = path.join(directory, `pdf-${round}`); await mkdir(output);
  const result = await run(executable, [profile, '--headless', '--convert-to', 'pdf:impress_pdf_Export', '--outdir', output, ...sets[round]], opts);
  logs.push({ phase: `pdf-${round}`, stdout: result.stdout, stderr: result.stderr });
  const files = ['control.pdf', 'reordered.pdf'].map(name => path.join(output, name));
  await Promise.all(files.map(file => readFile(file))); pdfSets.push(files);
}
const hash = bytes => createHash('sha256').update(bytes).digest('hex');
const files = await Promise.all([...sets.flat(), ...pdfSets.flat()].map(async file => ({ file: path.relative(process.cwd(), file), sha256: hash(await readFile(file)) })));
const independent = JSON.parse((await run('python3', ['scripts/research/reorder-libreoffice-oracle.py', directory], opts)).stdout);
const evidence = { schemaVersion: 1, scope: 'One overlapping top-level and two nested group reorderings, two fresh PPTX exports and three LibreOfficeDev headless saves/reopens, with independent XML and PDF observations. No interactive, repair-dialog, assistive-technology or PowerPoint claim.', environment: { application, platform: platform(), release: release(), arch: arch() }, directory: path.relative(process.cwd(), directory), runnerSha256: hash(await readFile('scripts/inspect-reorder-libreoffice.mjs')), oracleSha256: hash(await readFile('scripts/research/reorder-libreoffice-oracle.py')), files, logs, independent };
await writeFile('docs/testing/reorder-libreoffice-evidence.json', JSON.stringify(evidence, null, 2) + '\n');
console.log(JSON.stringify({ directory: evidence.directory, application, independent }, null, 2));
