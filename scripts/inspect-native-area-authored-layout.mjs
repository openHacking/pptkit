import { readFile,writeFile } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { importPptx } from '../packages/pptx-parser/dist/index.js';
import { resolveLayout } from '../packages/layout/dist/index.js';
const dir='.tmp/native-area-authored/powerpoint-ohuZBk',bytes=await readFile(`${dir}/area-authored.pptx`), session=await importPptx(bytes,{projection:'content'});
try { const slides=resolveLayout(session.document).slides.map(s=>{const e=s.elements.find(e=>e.type==='chart');return {box:e.box,grouping:e.grouping,blankDisplay:e.blankDisplay,plotBox:e.chartLayout.plotBox,valueScale:e.chartLayout.valueScale,areas:e.chartLayout.areas};});await writeFile(`${dir}/layout.json`,JSON.stringify({sourceSha256:createHash('sha256').update(bytes).digest('hex'),slides},null,2)+'\n'); } finally {session.dispose();}
