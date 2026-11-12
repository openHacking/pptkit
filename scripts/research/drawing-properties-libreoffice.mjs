// Named producer experiment: source-preserving edits, then three LibreOffice headless saves.
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { spawnSync } from 'node:child_process';
import { mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises';
import { platform, release, arch } from 'node:os';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { createPresentation } from '../../packages/core/dist/index.js';
import { generatePptx } from '../../packages/pptx-exporter/dist/index.js';
import { importPptx } from '../../packages/pptx-parser/dist/index.js';
import { readGeneratedZip } from '../../tests/support/read-generated-zip.mjs';

const root = process.cwd(), base = path.join(root, '.tmp', 'drawing-properties-libreoffice');
await mkdir(base, { recursive: true }); const dir = await mkdtemp(path.join(base, 'run-'));
const rel = file => path.relative(root, file).split(path.sep).join('/');
const sha = value => createHash('sha256').update(value).digest('hex');
function run(command, args) { const result = spawnSync(command, args, { encoding: 'utf8', timeout: 90000 }); assert.equal(result.status, 0, `${command}: ${result.error?.message ?? result.stderr}`); return { stdout: result.stdout.trim(), stderr: result.stderr.trim() }; }
const oracle = path.join(root, 'scripts/research/drawing-properties-oracle.py');
const inspect = file => JSON.parse(run('python3', [oracle, file]).stdout);
const deck = createPresentation({ size: { width: 400, height: 225 } });
const text = (id, name, description, label) => ({ type: 'text', id, name, box: { x: 30, y: id === 'target' ? 25 : 105, width: 300, height: 45 }, content: label, accessibility: { description, decorative: true } });
deck.addSlide({ id: 'first', elements: [text('target', 'Original target', 'Original description', 'Original text'), text('sibling', 'Sibling', 'Sibling description', 'Sibling text')] });
deck.addSlide({ id: 'second', elements: [text('other', 'Other page', 'Other description', 'Other text')] });
const generated = await generatePptx(deck); assert.deepEqual(generated.warnings, []);
const originalPath = path.join(dir, 'original.pptx'); await writeFile(originalPath, generated.bytes);
const session = await importPptx(generated.bytes, { projection: 'content', editSaveProfile: 'text-runs-v3', descriptionEditSaveProfile: 'mixed-v5', drawingNameSaveProfile: 'existing-v1', decorativeSaveProfile: 'existing-v1', firstSlideNumberSaveProfile: 'attribute-v2' });
let edited;
try {
  const slide = session.document.slides[0], original = slide.elements[0], next = structuredClone(original);
  next.name = 'Edited target 中文'; next.accessibility.description = 'Edited description & 中文'; next.accessibility.decorative = false; next.content[0].runs[0].text = 'Edited text'; slide.replaceElement(original.id, next); session.document.setFirstSlideNumber(17);
  const plan = await session.prepareSave({ mode: 'edit' }); assert.equal(plan.status, 'ready', JSON.stringify(plan.report)); assert.equal(plan.plan.report.changes.length, 5);
  edited = (await session.save(plan.plan)).bytes;
} finally { session.dispose(); }
const sourcePath = path.join(dir, 'edited.pptx'); await writeFile(sourcePath, edited);
const before = readGeneratedZip(generated.bytes), after = readGeneratedZip(edited), changedParts = [];
assert.deepEqual([...before.keys()], [...after.keys()]);
for (const [name, bytes] of before) if (!Buffer.from(bytes).equals(Buffer.from(after.get(name)))) changedParts.push(name);
assert.deepEqual(changedParts.sort(), ['ppt/presentation.xml', 'ppt/slides/slide1.xml']);
const baseline = inspect(sourcePath); assert.equal(baseline.firstSlideNumber, '17'); assert.equal(baseline.pages.length, 2);
const target = baseline.pages[0].objects.find(x => x.text === 'Edited text'), sibling = baseline.pages[0].objects.find(x => x.text === 'Sibling text'), other = baseline.pages[1].objects.find(x => x.text === 'Other text');
assert.deepEqual(target, { name: 'Edited target 中文', description: 'Edited description & 中文', decorativeValues: ['0'], text: 'Edited text' });
assert.deepEqual(sibling, { name: 'Sibling', description: 'Sibling description', decorativeValues: ['1'], text: 'Sibling text' });
assert.deepEqual(other, { name: 'Other page', description: 'Other description', decorativeValues: ['1'], text: 'Other text' });
const version = run('soffice', ['--version']).stdout, logs = [], rounds = [{ round: 0, path: rel(sourcePath), sha256: sha(edited), inventory: baseline }];
let source = sourcePath;
for (let round = 1; round <= 3; round++) {
  const out = path.join(dir, `round-${round}`); await mkdir(out);
  const profile = path.join(dir, `profile-${round}`);
  const log = run('soffice', [`-env:UserInstallation=${pathToFileURL(profile).href}`, '--headless', '--convert-to', 'pptx', '--outdir', out, source]); logs.push({ round, ...log });
  source = path.join(out, 'edited.pptx'); const bytes = await readFile(source), inventory = inspect(source);
  assert.equal(inventory.pages.length, 2); assert.ok(inventory.pages[0].objects.some(x => x.text === 'Edited text')); assert.ok(inventory.pages[0].objects.some(x => x.text === 'Sibling text')); assert.ok(inventory.pages[1].objects.some(x => x.text === 'Other text'));
  rounds.push({ round, path: rel(source), sha256: sha(bytes), inventory });
}
const evidence = { schemaVersion: 1, scope: 'One two-slide PPTKit source-preservation transaction followed by three LibreOffice headless conversions; no interactive edit, repair-dialog, PowerPoint or assistive-technology observation.', environment: { application: version, platform: platform(), release: release(), arch: arch() }, runnerSha256: sha(await readFile(new URL(import.meta.url))), oracleSha256: sha(await readFile(oracle)), original: { path: rel(originalPath), sha256: sha(generated.bytes) }, sourceChangedParts: changedParts, logs, rounds };
await writeFile(path.join(root, 'docs/testing/drawing-properties-libreoffice-evidence.json'), JSON.stringify(evidence, null, 2) + '\n');
console.log(JSON.stringify({ directory: rel(dir), application: version, rounds: rounds.map(({ round, inventory }) => ({ round, inventory })) }, null, 2));
