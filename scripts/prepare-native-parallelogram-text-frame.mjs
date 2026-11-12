import { mkdir, writeFile } from "node:fs/promises";
import { createPresentation } from "../packages/core/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { renderPresentationToSvg } from "../packages/svg-renderer/dist/index.js";
const deck=createPresentation({size:{width:600,height:360}}), slide=deck.addSlide({id:"s"});
for(const [id,box,slant] of [["zero",{x:30,y:30,width:240,height:80},0],["mid",{x:30,y:140,width:240,height:80},.4],["tall",{x:330,y:30,width:110,height:240},.4]]) slide.addElement({id,name:id,type:"shape",shape:"parallelogram",slant,box,style:{fill:{type:"solid",color:"EAEAEA"},stroke:{paint:{type:"none"}}},text:{frame:{margin:0,wrap:false,verticalAlign:"top"},content:[{runs:[{text:"MMMM",style:{fontFamily:"Arial",fontSize:24,color:"000000"}}]}]}});
await mkdir(".tmp/native-parallelogram-text",{recursive:true});
await writeFile(".tmp/native-parallelogram-text/source.pptx",(await generatePptx(deck)).bytes,{flag:"wx"});
await writeFile(".tmp/native-parallelogram-text/source.svg",(await renderPresentationToSvg(deck)).slides[0].svg,{flag:"wx"});
