import assert from 'node:assert/strict';
import { readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { importPptx } from '../packages/pptx-parser/dist/index.js';
const directory=process.argv[2];if(!directory)throw Error('Usage: node scripts/prepare-native-v2-post-edit.mjs directory');
const session=await importPptx(await readFile(path.join(directory,'round3.pptx')),{projection:'content',editSaveProfile:'text-runs-v2'});
try {
 const slide=session.document.slides[0],element=slide.elements.find(item=>item.type==='shape' && item.text && typeof item.text.content!=='string' && item.text.content[0]?.runs[0]?.text==='Edited preserved shape 123 <&>');assert.ok(element);
 const updated=structuredClone(element),notes=structuredClone(slide.notes);assert.equal(notes[0].runs[0].text,'Edited preserved notes 456 <&>');updated.text.content[0].runs[0].text='PPTKit shape after UI <&>';notes[0].runs[0].text='PPTKit notes after UI <&>';slide.replaceElement(element.id,updated);slide.setNotes(notes);
 const result=await session.prepareSave({mode:'edit'});assert.equal(result.status,'ready',JSON.stringify(result.report));const saved=await session.save(result.plan);await writeFile(path.join(directory,'post-native-pptkit.pptx'),saved.bytes);await writeFile(path.join(directory,'post-native-report.json'),JSON.stringify(saved.report,null,2)+'\n');
} finally {session.dispose();}
