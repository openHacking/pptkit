import { readFile } from 'node:fs/promises';
import { importPptx } from '@pptkit/pptx-parser';
import { importPptxFile } from '@pptkit/pptx-parser/node';

const source = await readFile('picture-source.pptx'), image = await readFile('picture-replacement.gif');
const imported = await importPptxFile('picture-source.pptx', { projection: 'content', pictureEditSaveProfile: 'replace-v1' });
const session = imported.session;
try {
  const slide = session.document.slides[0], picture = slide.elements[0];
  session.document.registerAsset({ id: 'replacement', kind: 'image', source: { type: 'bytes', value: image }, mimeType: 'image/gif', width: 1, height: 1 });
  slide.replaceElement(picture.id, { ...picture, assetId: 'replacement' });
  const prepared = await session.prepareSave({ mode: 'edit' });
  if (prepared.status !== 'ready') throw new Error(JSON.stringify(prepared.report));
  const result = await session.save(prepared.plan), reopened = await importPptx(result.bytes, { projection: 'content', pictureSourceProfile: 'embedded-v1' });
  try { process.stdout.write(JSON.stringify({ sourceBytes: source.length, assurance: result.report.assurance, media: reopened.report.pictureSources.map(item => item.mediaPart) })); }
  finally { reopened.dispose(); }
} finally { session.dispose(); }
