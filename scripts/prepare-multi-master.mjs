import { chromium } from '@playwright/test';
import assert from 'node:assert/strict';
import { mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { createHash } from 'node:crypto';
import { createPresentation, normalizePresentation } from '../packages/core/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';
import { renderPresentationToSvg } from '../packages/svg-renderer/dist/index.js';
import { multiMasterDocument } from '../tests/fixtures/multi-master.mjs';
await mkdir('.tmp/multi-master', { recursive: true });
const directory = path.resolve(await mkdtemp('.tmp/multi-master/run-')), document = multiMasterDocument(createPresentation);
const generated = await generatePptx(document); assert.deepEqual(generated.warnings, []);
const file = path.join(directory, 'source.pptx'); await writeFile(file, generated.bytes);
const normalized = normalizePresentation(document); await writeFile(path.join(directory, 'source.json'), JSON.stringify(normalized, null, 2) + '\n');
const svg = await renderPresentationToSvg(document);
for (const [i, slide] of svg.slides.entries()) await writeFile(path.join(directory, `page-${i + 1}.svg`), slide.svg);
const browser = await chromium.launch();
const previewFile = path.join(directory, 'browser-preview.png'); let browserVersion;
try {
  browserVersion = browser.version();
  const page = await browser.newPage({ viewport: { width: 1000, height: 620 }, deviceScaleFactor: 1 });
  await page.route('**/*', route => route.abort());
  await page.setContent('<style>body{margin:16px;background:#ddd;display:grid;grid-template-columns:480px 480px;gap:12px}svg{width:480px;height:270px;background:white}</style>' + svg.slides.map(s => s.svg).join(''));
  await page.evaluate(() => document.fonts.ready);
  await page.screenshot({ path: previewFile });
} finally { await browser.close(); }
const artifacts = await Promise.all(['source.pptx', 'source.json', ...svg.slides.map((_, i) => `page-${i + 1}.svg`), 'browser-preview.png'].map(async name => {
  const file = path.join(directory, name), bytes = await readFile(file);
  return { file, bytes: bytes.length, sha256: createHash('sha256').update(bytes).digest('hex') };
}));
const manifest = { artifacts, schemaVersion: 1, recordedAt: new Date().toISOString(), directory, file, sha256: createHash('sha256').update(generated.bytes).digest('hex'), slides: normalized.slides.map((s, i) => ({ index: i + 1, id: s.id, layoutId: s.layoutId, masterId: normalized.layouts.find(l => l.id === s.layoutId).masterId ?? null, backgroundSource: s.backgroundSource })), browserPreview: { file: previewFile, sha256: createHash('sha256').update(await readFile(previewFile)).digest('hex'), browserVersion }, nativeReopen: 'not performed', sourceOnlyInspection: 'scripts/multi-master-parity.test.mjs', scope: 'Public authoring source, SVG and native OOXML graph; no native app save/edit/reopen or artwork admission yet.' };
await writeFile('docs/testing/multi-master-authoring-evidence.json', JSON.stringify(manifest, null, 2) + '\n'); console.log(JSON.stringify({ directory, sha256: manifest.sha256 }));
