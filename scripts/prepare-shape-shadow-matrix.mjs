import { writeFile } from 'node:fs/promises';
import { createPresentation } from '../packages/core/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';

// 25 separate pages avoid overlap or blending between test cases.
const output=process.argv[2];if(!output)throw Error('Expected new output path');
const d=createPresentation({size:{width:300,height:200}}),values=[0,.25,.5,.75,1];
for(const [row,shapeOpacity] of values.entries())for(const [column,shadowOpacity] of values.entries())d.addSlide({id:`shape-${row}-shadow-${column}`}).addElement({id:`rect-${row}-${column}`,type:'shape',shape:'rect',box:{x:40,y:50,width:80,height:80},opacity:shapeOpacity,style:{fill:{type:'solid',color:'FF0000'},stroke:{paint:{type:'none'}},shadow:{color:'000000',opacity:shadowOpacity,blur:0,distance:100,direction:0}}});
await writeFile(output,(await generatePptx(d)).bytes,{flag:'wx'});
