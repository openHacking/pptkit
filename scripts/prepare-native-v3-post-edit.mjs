import assert from 'node:assert/strict';
import { readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { importPptx } from '../packages/pptx-parser/dist/index.js';
const directory=process.argv[2];if(!directory)throw Error('Usage: node scripts/prepare-native-v3-post-edit.mjs directory');
const session=await importPptx(await readFile(path.join(directory,'round3.pptx')),{projection:'content',editSaveProfile:'text-runs-v3'});
try {
 const slide=session.document.slides[0],table=slide.elements.find(item=>item.type==='table' && Array.isArray(item.rows[0]?.cells[0]?.content) && item.rows[0].cells[0].content[0]?.runs[0]?.text==='Anchor 123 UIA');assert.ok(table);
 const updated=structuredClone(table),anchor=updated.rows[0].cells[0],right=updated.rows[1].cells[0];assert.equal(anchor.rowSpan,2);assert.equal(anchor.colSpan,2);assert.equal(right.content[0].runs[0].text,'Cell 456 UIB');anchor.content[0].runs[0].text='Anchor PPTKit UIA';right.content[0].runs[0].text='Cell PPTKit UIB';slide.replaceElement(table.id,updated);
 const result=await session.prepareSave({mode:'edit'});assert.equal(result.status,'ready',JSON.stringify(result.report));const saved=await session.save(result.plan);await writeFile(path.join(directory,'post-native-pptkit.pptx'),saved.bytes);await writeFile(path.join(directory,'post-native-report.json'),JSON.stringify(saved.report,null,2)+'\n');
}finally{session.dispose()}
