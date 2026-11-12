import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { chromium } from '@playwright/test';
import { createPresentation, normalizePresentation, resolvePresentationInheritance } from '../packages/core/dist/index.js';
import { resolveNormalizedLayout } from '../packages/layout/dist/index.js';
import { renderLayoutToSvg } from '../packages/svg-renderer/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';
import { chartNotesOriginsDocument } from '../tests/fixtures/chart-notes-origins.mjs';
import { textOriginsDocument } from '../tests/fixtures/text-origins.mjs';
const sha = bytes => createHash('sha256').update(bytes).digest('hex');
await mkdir('.tmp/chart-notes-origins', { recursive: true });
const directory = path.resolve(await mkdtemp('.tmp/chart-notes-origins/run-'));
const d = chartNotesOriginsDocument(createPresentation), resolution = resolvePresentationInheritance(d);
assert.deepEqual(resolution.presentation, normalizePresentation(d));
const layout = resolveNormalizedLayout(resolution.presentation), svg = await renderLayoutToSvg(layout), output = await generatePptx(d);
const names = [];
async function save(name, value) { names.push(name); await writeFile(path.join(directory, name), typeof value === 'string' || value instanceof Uint8Array ? value : JSON.stringify(value, null, 2) + '\n'); }
await save('source.pptx', output.bytes); await save('normalized.json', resolution.presentation); await save('origins.json', resolution.origins); await save('layout.json', layout);
for (const [index, slide] of svg.slides.entries()) await save(`page-${index+1}.svg`, slide.svg);
const old = textOriginsDocument(createPresentation), group = old.slides[0].elements.find(e => e.id === 'group');
const { placeholderKey: _binding, ...ordinaryGroup } = group; old.slides[0].replaceElement(group.id, ordinaryGroup);
const retainedSource = (await generatePptx(old)).bytes, expectedSha = 'ef3d886dcf09d80e23b58cbbafb67cbb26be84fbb2a5c71797c2eb0a3219cd2d';
assert.equal(sha(retainedSource), expectedSha); await save('retained-native-source.pptx', retainedSource);
const browser = await chromium.launch(); let browserVersion;
try {
  browserVersion = browser.version(); const page = await browser.newPage({ viewport: { width: 1000, height: 620 }, deviceScaleFactor: 1 });
  await page.route('**/*', r => r.abort());
  await page.setContent('<style>body{margin:8px;display:grid;grid-template-columns:480px 480px;gap:8px;background:#ddd}svg{width:480px;height:270px;background:white}</style>' + svg.slides.map(s => s.svg).join(''));
  await page.evaluate(() => document.fonts.ready); const name = 'browser-preview.png'; names.push(name); await page.screenshot({ path: path.join(directory, name) });
} finally { await browser.close(); }
const artifacts = await Promise.all(names.map(async name => { const file = path.join(directory, name), bytes = await readFile(file); return { file, bytes: bytes.length, sha256: sha(bytes) }; }));
const dependencies = await Promise.all(['tests/fixtures/chart-notes-origins.mjs', 'packages/core/src/normalization/presentation-normalizer.ts', 'packages/core/src/normalization/text-origins.ts', 'packages/core/src/types/inheritance.ts', 'packages/core/test/chart-notes-origins.test.mjs', 'scripts/multi-master-parity.test.mjs', 'scripts/prepare-chart-notes-origins.mjs'].map(async file => ({ file, sha256: sha(await readFile(file)) })));
const manifest = { schemaVersion: 1, recordedAt: new Date().toISOString(), directory, artifacts, dependencies, browserVersion, originRecords: resolution.origins.length, chartCount: 11, notesTargets: d.slides.map(s => s.id), warnings: output.warnings, retainedNativeSourceSha256: expectedSha, nativeApplication: 'No new application run; this slice changes origin metadata only. The retained two-page native source stays byte-exact. Chart/notes output is checked independently in browser XML tests; full native chart/theme admission remains open.' };
await writeFile('docs/testing/chart-notes-origins-evidence.json', JSON.stringify(manifest, null, 2)+'\n');
console.log(JSON.stringify({ directory, artifactCount: artifacts.length, records: resolution.origins.length, warnings: output.warnings.length, retainedNativeSourceSha256: expectedSha }));
