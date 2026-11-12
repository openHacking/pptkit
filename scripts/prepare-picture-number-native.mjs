import { mkdir, readFile, writeFile, copyFile } from 'node:fs/promises';
import { createPresentation } from '../packages/core/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';
import { importPptx } from '../packages/pptx-parser/dist/index.js';
import { png } from '../packages/pptx-parser/test/image-fixture.mjs';

const directory = process.argv[2]; if (!directory) throw new Error('Supply a new output directory.');
await mkdir(directory, { recursive: false });
const replacement = await readFile('tests/fixtures/transparent-replacement.png'), doc = createPresentation({ size: { width: 500, height: 250 } });
doc.registerAsset({ id: 'shared', kind: 'image', source: { type: 'bytes', value: png }, mimeType: 'image/png', width: 8, height: 4 });
doc.addSlide({ id: 'picture-slide', elements: [
  { id: 'target', type: 'image', assetId: 'shared', fit: 'crop', crop: { left: 0.125 }, box: { x: 30, y: 30, width: 150, height: 150 } },
  { id: 'sibling', type: 'image', assetId: 'shared', fit: 'crop', crop: { left: 0.125 }, box: { x: 300, y: 30, width: 150, height: 150 } },
] });
doc.addSlide({ id: 'control', elements: [{ id: 'control-title', type: 'text', box: { x: 30, y: 30, width: 400, height: 60 }, content: 'Control page' }] });
const source = (await generatePptx(doc)).bytes; await writeFile(`${directory}/source.pptx`, source, { flag: 'wx' });
const session = await importPptx(source, { projection: 'content', pictureEditSaveProfile: 'replace-v1', firstSlideNumberSaveProfile: 'attribute-v2' });
try {
  const slide = session.document.slides[0], target = slide.elements[0];
  session.document.registerAsset({ id: 'replacement', kind: 'image', source: { type: 'bytes', value: replacement }, mimeType: 'image/png', width: 4, height: 2 });
  slide.replaceElement(target.id, { ...target, assetId: 'replacement' }); session.document.setFirstSlideNumber(17);
  const prepared = await session.prepareSave({ mode: 'edit' }); if (prepared.status !== 'ready' || prepared.plan.report.changes.length !== 3) throw Error(JSON.stringify(prepared.report));
  await writeFile(`${directory}/candidate.pptx`, (await session.save(prepared.plan)).bytes, { flag: 'wx' }); await copyFile(`${directory}/candidate.pptx`, `${directory}/native-copy.pptx`);
} finally { session.dispose(); }
