import assert from "node:assert/strict";
import { createRequire } from "node:module";
import { writeFile } from "node:fs/promises";
import { createHash } from "node:crypto";
import { createPresentation, normalizePresentation } from "../packages/core/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { importPptx } from "../packages/pptx-parser/dist/index.js";

const dir=process.argv[2];if(!dir||process.argv.length!==3)throw Error("Expected an existing empty evidence directory.");
const require=createRequire(new URL("../packages/pptx-parser/package.json",import.meta.url));const {unzipSync,zipSync,strFromU8,strToU8}=require("fflate");
const d=createPresentation();d.defineSlideMaster({id:"matrix-master",name:"Matrix master",background:{type:"solid",color:"A00000"}});
d.defineSlideLayout({id:"master-bg",name:"Inherited",masterId:"matrix-master"});d.defineSlideLayout({id:"layout-bg",name:"Layout background",masterId:"matrix-master",background:{type:"solid",color:"008000"}});
for(const [i,layout,background] of [[1,"master-bg",undefined],[2,"layout-bg",undefined],[3,"layout-bg",{type:"solid",color:"0000C0"}]]){
 const slide=d.addSlide({id:`page-${i}`,layoutId:layout,...(background?{background}:{})});slide.addElement({id:`title-${i}`,type:"text",box:{x:40,y:35,width:400,height:60},content:[{runs:[{text:`BGREF PAGE ${i}`,style:{fontFamily:"Arial",fontSize:28,color:"FFFFFF"}}]}]});
}
const generated=(await generatePptx(d)).bytes,files=unzipSync(generated);let replacements=0;
for(const [name,bytes] of Object.entries(files)){
 if(!/^ppt\/(slideMasters|slideLayouts|slides)\/[^/]+\.xml$/.test(name))continue;
 let xml=strFromU8(bytes);xml=xml.replace(/<p:bg><p:bgPr><a:solidFill><a:srgbClr val="(A00000|008000|0000C0)"><\/a:srgbClr><\/a:solidFill><a:effectLst\/><\/p:bgPr><\/p:bg>/g,(_all,color)=>{replacements++;return `<p:bg><p:bgRef idx="1001"><a:srgbClr val="${color}"/></p:bgRef></p:bg>`});files[name]=strToU8(xml);
}
assert.equal(replacements,3,"expected one master, one layout and one slide background");const source=zipSync(files),session=await importPptx(source,{projection:"content",masterArtwork:"artwork-v6"});
assert.deepEqual(normalizePresentation(session.document).slides.map(s=>s.background.color),["A00000","008000","0000C0"]);
const projected=(await generatePptx(session.document)).bytes;await writeFile(`${dir}/source-bgref.pptx`,source,{flag:"wx"});await writeFile(`${dir}/projected-bgref.pptx`,projected,{flag:"wx"});
const hash=b=>createHash("sha256").update(b).digest("hex");console.log(JSON.stringify({sourceSha256:hash(source),projectedSha256:hash(projected),replacements,backgrounds:normalizePresentation(session.document).slides.map(s=>({paint:s.background,source:s.backgroundSource}))},null,2));session.dispose();
