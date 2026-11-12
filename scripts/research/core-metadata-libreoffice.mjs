// Source-preserved metadata edit versus a control, followed by three producer conversions.
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

const root = process.cwd(), base = path.join(root, '.tmp', 'core-metadata-libreoffice'); await mkdir(base, { recursive: true }); const dir = await mkdtemp(path.join(base, 'run-'));
const rel = file => path.relative(root, file).split(path.sep).join('/'), sha = bytes => createHash('sha256').update(bytes).digest('hex');
function run(command, args) { const result = spawnSync(command, args, { encoding: 'utf8', timeout: 90000 }); assert.equal(result.status, 0, `${command}: ${result.error?.message ?? result.stderr}`); return { stdout: result.stdout.trim(), stderr: result.stderr.trim() }; }
const oracle = path.join(root, 'scripts/research/core-metadata-oracle.py');
const inspect = file => JSON.parse(run('python3', [oracle, 'inspect', file]).stdout);
const deck = createPresentation({ metadata: { title: 'Original title', author: 'Sentinel Author', subject: 'Original subject', description: 'Original description', language: 'en-US', keywords: ['alpha', 'beta'], revision: 7 } });
deck.addSlide({ elements: [{ type: 'text', name: 'Control', box: { x: 20, y: 20, width: 300, height: 60 }, content: 'Metadata control slide' }] });
const generated = await generatePptx(deck); assert.deepEqual(generated.warnings, []);
const input = path.join(dir, 'input'); await mkdir(input); const generatedPath = path.join(dir, 'generated.pptx'), controlPath = path.join(input, 'control.pptx'), editedPath = path.join(input, 'edited.pptx');
await writeFile(generatedPath, generated.bytes); run('python3', [oracle, 'inject', generatedPath, controlPath]); const control = await readFile(controlPath);
const session = await importPptx(control, { projection: 'content', metadataSaveProfile: 'core-strings-v1' }); let edited;
try { session.document.setMetadata({ ...session.document.metadata, title: 'Edited title 中文', subject: 'Edited subject', description: 'Edited & description', language: 'fr-FR' }); const plan = await session.prepareSave({ mode: 'edit' }); assert.equal(plan.status, 'ready', JSON.stringify(plan.report)); assert.equal(plan.plan.report.changes.length, 4); edited = (await session.save(plan.plan)).bytes; } finally { session.dispose(); }
await writeFile(editedPath, edited); const before = readGeneratedZip(control), after = readGeneratedZip(edited), changedParts = []; assert.deepEqual([...before.keys()], [...after.keys()]); for (const [name, bytes] of before) if (!Buffer.from(bytes).equals(Buffer.from(after.get(name)))) changedParts.push(name); assert.deepEqual(changedParts, ['docProps/core.xml']);
const get = (report, name) => report.properties.find(p => p.expandedName.endsWith(`}${name}`))?.value;
const controlInventory = inspect(controlPath), editedInventory = inspect(editedPath);
for (const [name, value] of [['creator', 'Sentinel Author'], ['revision', '7'], ['lastModifiedBy', 'Sentinel Maintainer'], ['created', '2020-01-02T03:04:05Z'], ['modified', '2021-02-03T04:05:06Z']]) { assert.equal(get(controlInventory, name), value); assert.equal(get(editedInventory, name), value); }
for (const [name, value] of [['title', 'Edited title 中文'], ['subject', 'Edited subject'], ['description', 'Edited & description'], ['language', 'fr-FR']]) assert.equal(get(editedInventory, name), value);
const version = run('soffice', ['--version']).stdout, python = run('python3', ['--version']).stdout, paths = { control: controlPath, edited: editedPath }, records = [], logs = [];
for (let round = 0; round <= 3; round++) {
  if (round) { const output = path.join(dir, `round-${round}`); await mkdir(output); for (const kind of ['control', 'edited']) { const log = run('soffice', [`-env:UserInstallation=${pathToFileURL(path.join(dir, `profile-${round}-${kind}`)).href}`, '--headless', '--convert-to', 'pptx', '--outdir', output, paths[kind]]); logs.push({ round, kind, ...log }); paths[kind] = path.join(output, `${kind}.pptx`); } }
  for (const kind of ['control', 'edited']) { const bytes = await readFile(paths[kind]), inventory = inspect(paths[kind]); assert.equal(inventory.slides.length, 1); assert.match(inventory.slides[0].text, /Metadata control slide/); records.push({ round, kind, path: rel(paths[kind]), sha256: sha(bytes), inventory }); }
}
const evidence = { schemaVersion: 1, scope: 'One PPTKit core string edit and unchanged control, each resaved three times by LibreOfficeDev headlessly. XML inventories do not establish native metadata UI or interactive behavior.', environment: { application: version, platform: platform(), release: release(), arch: arch(), python }, runnerSha256: sha(await readFile(new URL(import.meta.url))), oracleSha256: sha(await readFile(oracle)), generated: { path: rel(generatedPath), sha256: sha(generated.bytes) }, sourceChangedParts: changedParts, logs, records };
await writeFile(path.join(root, 'docs/testing/core-metadata-libreoffice-evidence.json'), JSON.stringify(evidence, null, 2) + '\n'); console.log(JSON.stringify({ directory: rel(dir), application: version, records: records.map(({ round, kind, inventory }) => ({ round, kind, properties: inventory.properties })) }, null, 2));
