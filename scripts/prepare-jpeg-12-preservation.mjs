import { readFile, writeFile } from 'node:fs/promises';
import { importPptx } from '../packages/pptx-parser/dist/index.js';

const destination = process.argv[2] ?? '.tmp/native-jpeg-12/preserved-before.pptx';
const source = await readFile('.tmp/native-jpeg-12/probe.pptx');
const image = await readFile('tests/fixtures/jpeg-12bit-progressive.jpg');
const session = await importPptx(source, { projection: 'content', pictureEditSaveProfile: 'replace-v1' });
try {
  const slide = session.document.slides[0], pictures = slide.elements.filter(element => element.type === 'image');
  if (pictures.length !== 3) throw new Error('Expected three source pictures.');
  session.document.registerAsset({ id: 'replacement12', kind: 'image', source: { type: 'bytes', value: image }, mimeType: 'image/jpeg', width: 3, height: 2 });
  slide.replaceElement(pictures[2].id, { ...pictures[2], assetId: 'replacement12' });
  const prepared = await session.prepareSave({ mode: 'edit' });
  if (prepared.status !== 'ready') throw new Error(JSON.stringify(prepared.report));
  await writeFile(destination, (await session.save(prepared.plan)).bytes);
} finally { session.dispose(); }
