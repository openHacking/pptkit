import assert from 'node:assert/strict';
import { readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { importPptx } from '../packages/pptx-parser/dist/index.js';
const directory=process.argv[2];if(!directory)throw Error('Usage: node scripts/prepare-native-v1-post-edit.mjs directory');
const session=await importPptx(await readFile(path.join(directory,'round3.pptx')),{projection:'content',editSaveProfile:'text-runs-v1'});
try {
 const slide=session.document.slides[0],element=slide.elements.find(item=>item.type==='text' && typeof item.content!=='string' && item.content[0]?.runs[0]?.text==='Edited preserved text 123 <&>');assert.ok(element);
 const updated=structuredClone(element);updated.content[0].runs[0].text='PPTKit after native 456 <&>';const index=slide.elements.indexOf(element);slide.replaceElement(element.id,updated);
 const result=await session.prepareSave({mode:'edit'});assert.equal(result.status,'ready',JSON.stringify(result.report));const saved=await session.save(result.plan);await writeFile(path.join(directory,'post-native-pptkit.pptx'),saved.bytes);await writeFile(path.join(directory,'post-native-report.json'),JSON.stringify(saved.report,null,2)+'\n');assert.equal(index,slide.elements.indexOf(slide.elements.find(item=>item.id===element.id)));
} finally {session.dispose();}
