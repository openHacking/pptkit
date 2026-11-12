import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { chromium } from '@playwright/test';
import { createPresentation, resolvePresentationInheritance } from '../packages/core/dist/index.js';
import { resolveNormalizedLayout } from '../packages/layout/dist/index.js';
import { renderLayoutToSvg } from '../packages/svg-renderer/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';
import { textOriginsDocument } from '../tests/fixtures/text-origins.mjs';
import { multiMasterDocument } from '../tests/fixtures/multi-master.mjs';
import { readGeneratedZip } from '../tests/support/read-generated-zip.mjs';

const sha = bytes => createHash('sha256').update(bytes).digest('hex');
await mkdir('.tmp/text-origins', { recursive: true });
const directory = path.resolve(await mkdtemp('.tmp/text-origins/run-'));
const d = textOriginsDocument(createPresentation), group = d.slides[0].elements.find(e => e.id === 'group');
const { placeholderKey: _binding, ...ordinaryGroup } = group;
d.slides[0].replaceElement(group.id, ordinaryGroup); // Native candidate excludes the unadmitted group-placeholder binding.
const resolution = resolvePresentationInheritance(d), layout = resolveNormalizedLayout(resolution.presentation), svg = await renderLayoutToSvg(layout), generated = await generatePptx(d);
const output = async (name, value) => writeFile(path.join(directory, name), typeof value === 'string' || value instanceof Uint8Array ? value : JSON.stringify(value, null, 2) + '\n');
await output('source.pptx', generated.bytes); await output('normalized.json', resolution.presentation); await output('origins.json', resolution.origins); await output('layout.json', layout);
for (const [i, slide] of svg.slides.entries()) await output(`page-${i + 1}.svg`, slide.svg);
const oldFile = path.resolve('.tmp/multi-master/native-b85qrfu0/fixed.pptx'), old = await readFile(oldFile);
assert.equal(sha(old), '943270e78b2fe373a20e2983c239b22211ee4c973bf9fa9d23f5b64fa0576e89');
const updated = (await generatePptx(multiMasterDocument(createPresentation))).bytes;
await output('legacy-current.pptx', updated);
const a = readGeneratedZip(old), b = readGeneratedZip(updated);
assert.deepEqual([...a.keys()].sort(), [...b.keys()].sort());
const changes = [];
for (const [name, bytes] of a) {
  const current = b.get(name); if (bytes.equals(current)) continue;
  let reduced = current.toString(), kind;
  if (/^ppt\/(slides\/slide|notesSlides\/notesSlide)[1-4]\.xml$/.test(name)) {
    kind = 'explicit-false-run-flags';
    reduced = reduced.replace(/<(a:(?:rPr|endParaRPr))\b[^>]*>/g, tag => tag.replace(/ (?:b="0"|i="0"|u="none"|strike="noStrike")/g, ''));
  } else if (/^ppt\/slideLayouts\/slideLayout[2-5]\.xml$/.test(name)) {
    kind = 'shared-placeholder-defaults';
    assert.equal((reduced.match(/<a:pPr\b/g) ?? []).length, 1); assert.equal((bytes.toString().match(/<a:pPr\b/g) ?? []).length, 0);
    reduced = reduced.replace(/<a:pPr\b[^>]*>[\s\S]*?<\/a:pPr>/g, '');
  } else throw new Error(`Unexpected legacy payload change: ${name}`);
  assert.equal(reduced, bytes.toString(), name);
  changes.push({ part: name, kind, before: sha(bytes), after: sha(current) });
}
assert.equal(changes.length, 12);
const browser = await chromium.launch(); let browserVersion;
try {
  browserVersion = browser.version();
  const page = await browser.newPage({ viewport: { width: 1000, height: 310 }, deviceScaleFactor: 1 });
  await page.route('**/*', r => r.abort());
  await page.setContent('<style>body{margin:8px;display:flex;gap:8px;background:#ddd}svg{width:480px;height:270px;background:white}</style>' + svg.slides.map(s => s.svg).join(''));
  await page.evaluate(() => document.fonts.ready); await page.screenshot({ path: path.join(directory, 'browser-preview.png') });
} finally { await browser.close(); }
const names = ['source.pptx', 'normalized.json', 'origins.json', 'layout.json', 'legacy-current.pptx', ...svg.slides.map((_, i) => `page-${i + 1}.svg`), 'browser-preview.png'];
const artifacts = await Promise.all(names.map(async name => { const file = path.join(directory, name), bytes = await readFile(file); return { file, bytes: bytes.length, sha256: sha(bytes) }; }));
const manifest = { schemaVersion: 1, recordedAt: new Date().toISOString(), directory, artifacts, sourceFile: path.join(directory, 'source.pptx'), warnings: generated.warnings, originRecords: resolution.origins.length, legacy: { originalFile: oldFile, originalSha256: sha(old), currentSha256: sha(updated), changes }, browserVersion, nativeReopen: 'not performed', scope: 'Public authoring candidate with text format sources, optional-font clear barriers, shared placeholder defaults and explicit false flags. Parent group-placeholder binding deliberately removed in this native candidate. No native application or live template-edit admission.' };
try {
  const initial = JSON.parse(await readFile('docs/testing/text-origins-initial-failure-evidence.json', 'utf8'));
  const previous = await readFile(initial.sourceFile), oldParts = readGeneratedZip(previous), newParts = readGeneratedZip(generated.bytes);
  assert.deepEqual([...oldParts.keys()].sort(), [...newParts.keys()].sort());
  const changed = [...oldParts].filter(([name, value]) => !value.equals(newParts.get(name)));
  assert.deepEqual(changed.map(([name]) => name), ['ppt/slideLayouts/slideLayout2.xml']);
  const before = changed[0][1].toString(), artwork = [...before.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g)].find(m => m[0].includes('name="shared-text"'))[0];
  assert.equal((artwork.match(/<p:ph\b/g) ?? []).length, 1);
  assert.equal(before.replace(artwork, artwork.replace(/<p:ph\b[^>]*\/>/, '')), newParts.get(changed[0][0]).toString());
  manifest.nativeFailureDelta = { originalFile: initial.sourceFile, originalSha256: sha(previous), changedPart: changed[0][0], change: 'Remove only the extra native placeholder identity from the shared layout object; every other payload is exact.' };
} catch (error) { if (error.code !== 'ENOENT') throw error; }
await writeFile('docs/testing/text-origins-evidence.json', JSON.stringify(manifest, null, 2) + '\n');
console.log(JSON.stringify({ directory, sourceFile: manifest.sourceFile, records: manifest.originRecords, changes: changes.length, warnings: generated.warnings.length }));
