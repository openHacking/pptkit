// Source-preserving visibility edit followed by three independent headless producer rounds.
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

const root = process.cwd(), base = path.join(root, '.tmp', 'slide-visibility-libreoffice'); await mkdir(base, { recursive: true });
const dir = await mkdtemp(path.join(base, 'run-')), rel = file => path.relative(root, file).split(path.sep).join('/'), sha = bytes => createHash('sha256').update(bytes).digest('hex');
function run(command, args) { const result = spawnSync(command, args, { encoding: 'utf8', timeout: 90000 }); assert.equal(result.status, 0, `${command}: ${result.error?.message ?? result.stderr}`); return { stdout: result.stdout.trim(), stderr: result.stderr.trim() }; }
const oracle = path.join(root, 'scripts/research/slide-visibility-oracle.py'), deck = createPresentation({ size: { width: 400, height: 225 } });
for (const [id, label] of [['one', 'Visible Alpha'], ['two', 'Target Beta'], ['three', 'Visible Gamma']]) deck.addSlide({ id, elements: [{ type: 'text', name: label, box: { x: 30, y: 45, width: 300, height: 60 }, content: label }] });
const fresh = await generatePptx(deck); assert.deepEqual(fresh.warnings, []);
const originalDir = path.join(dir, 'input'); await mkdir(originalDir); const controlPath = path.join(originalDir, 'control.pptx'), hiddenPath = path.join(originalDir, 'hidden.pptx'); await writeFile(controlPath, fresh.bytes);
const session = await importPptx(fresh.bytes, { projection: 'content', slideVisibilitySaveProfile: 'attribute-v1' }); let hidden;
try { assert.equal(session.document.slides[1].hidden, false); session.document.slides[1].setHidden(true); const plan = await session.prepareSave({ mode: 'edit' }); assert.equal(plan.status, 'ready', JSON.stringify(plan.report)); assert.equal(plan.plan.report.changes.length, 1); hidden = (await session.save(plan.plan)).bytes; } finally { session.dispose(); }
await writeFile(hiddenPath, hidden);
const before = readGeneratedZip(fresh.bytes), after = readGeneratedZip(hidden), changedParts = []; assert.deepEqual([...before.keys()], [...after.keys()]);
for (const [name, bytes] of before) if (!Buffer.from(bytes).equals(Buffer.from(after.get(name)))) changedParts.push(name);
assert.deepEqual(changedParts, ['ppt/slides/slide2.xml']);
const logs = [], records = [], paths = { control: controlPath, hidden: hiddenPath }, version = run('soffice', ['--version']).stdout;
const python = JSON.parse(run('python3', ['-c', 'import fitz,json,sys; print(json.dumps({"python":sys.version.split()[0],"pymupdf":fitz.VersionBind}))']).stdout);
for (let round = 0; round <= 3; round++) {
  if (round) {
    const output = path.join(dir, `round-${round}`); await mkdir(output);
    for (const kind of ['control', 'hidden']) { const log = run('soffice', [`-env:UserInstallation=${pathToFileURL(path.join(dir, `profile-save-${round}-${kind}`)).href}`, '--headless', '--convert-to', 'pptx', '--outdir', output, paths[kind]]); logs.push({ phase: 'save', round, kind, ...log }); paths[kind] = path.join(output, `${kind}.pptx`); }
  }
  const output = path.join(dir, `pdf-${round}`); await mkdir(output);
  for (const kind of ['control', 'hidden']) {
    const log = run('soffice', [`-env:UserInstallation=${pathToFileURL(path.join(dir, `profile-pdf-${round}-${kind}`)).href}`, '--headless', '--convert-to', 'pdf', '--outdir', output, paths[kind]]); logs.push({ phase: 'pdf', round, kind, ...log });
    const pdf = path.join(output, `${kind}.pdf`), result = JSON.parse(run('python3', [oracle, paths[kind], pdf]).stdout);
    assert.equal(result.pptx.slides.length, 3);
    assert.deepEqual(result.pptx.slides.map(slide => slide.text), ['Visible Alpha', 'Target Beta', 'Visible Gamma']);
    if (round === 0) assert.deepEqual(result.pptx.slides.map(slide => slide.hidden), kind === 'control' ? [false, false, false] : [false, true, false]);
    records.push({ round, kind, pptx: { path: rel(paths[kind]), sha256: sha(await readFile(paths[kind])) }, pdf: { path: rel(pdf), sha256: sha(await readFile(pdf)) }, inventory: result });
  }
}
const evidence = { schemaVersion: 1, scope: 'One three-slide PPTKit visibility edit versus an unchanged control, three LibreOfficeDev headless PPTX resaves and four PDF exports per input. PDF extraction is not slideshow or interactive verification.', environment: { application: version, platform: platform(), release: release(), arch: arch(), ...python }, runnerSha256: sha(await readFile(new URL(import.meta.url))), oracleSha256: sha(await readFile(oracle)), sourceChangedParts: changedParts, logs, records };
await writeFile(path.join(root, 'docs/testing/slide-visibility-libreoffice-evidence.json'), JSON.stringify(evidence, null, 2) + '\n');
console.log(JSON.stringify({ directory: rel(dir), application: version, rounds: records.map(({ round, kind, inventory }) => ({ round, kind, visibility: inventory.pptx.slides.map(s => s.hidden), show: inventory.pptx.slides.map(s => s.show), pdf: inventory.pdf.pages.map(p => p.text.trim()) })) }, null, 2));
