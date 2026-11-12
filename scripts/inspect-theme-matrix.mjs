import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { mkdir,mkdtemp,readFile,writeFile } from "node:fs/promises";
import { createHash } from "node:crypto";
import { platform,release,arch } from "node:os";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { createPresentation,normalizePresentation } from "../packages/core/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";
import { zip } from "../packages/pptx-parser/test/zip-fixture.mjs";
import { matrixFixture,formatScheme,matrixFill,fillReference,lineReference } from "../packages/pptx-parser/test/theme-matrix-fixture.mjs";
import { TABLE_STYLE_ID } from "../packages/pptx-parser/test/table-style-fixture.mjs";
if(process.argv.length>3)throw new Error("Usage: node scripts/inspect-theme-matrix.mjs [soffice-executable]");
const execute=promisify(execFile),executable=process.argv[2]??"soffice",options={timeout:45000,maxBuffer:8*1024*1024},hash=bytes=>createHash("sha256").update(bytes).digest("hex"),application=(await execute(executable,["--version"],options)).stdout.trim();
await mkdir(".tmp/theme-matrix-native",{recursive:true});const directory=await mkdtemp(path.resolve(".tmp/theme-matrix-native/run-")),nativeDir=path.join(directory,"native"),pdfDir=path.join(directory,"pdf"),controlDir=path.join(directory,"source-pdf");for(const dir of[nativeDir,pdfDir,controlDir])await mkdir(dir);
const expected=[],sources=[],hashes=[],failures=[];
for(const name of["placeholder","background","no-fill","no-line","transforms"]){
  const document=createPresentation({size:{width:400,height:300},theme:{colors:{accent1:"3579BD",accent2:"CA6420"},fonts:{body:"Arial",heading:"Arial"}}});document.addSlide().addElement({type:"table",box:{x:30,y:30,width:280,height:120},columns:Array(4).fill(70),rows:Array.from({length:4},(_,r)=>({height:30,cells:Array.from({length:4},(_,c)=>({content:[{runs:[{text:`R${r}C${c}`,style:{fontFamily:"Arial",fontSize:10,color:"000000"}}]}]}))}))});
  let fill=fillReference(),border=lineReference(),scheme=formatScheme();
  if(name==="background"){fill=fillReference(1001);border=lineReference(3,'<a:srgbClr val="0000FF"/>');}
  if(name==="no-fill")fill=fillReference(1000);
  if(name==="no-line")border=lineReference(0);
  if(name==="transforms"){fill=fillReference(1,'<a:srgbClr val="000000"><a:tint val="99500"/><a:alpha val="40000"/></a:srgbClr>');scheme=formatScheme({fills:[matrixFill("phClr",'<a:shade val="50000"/><a:alphaMod val="50000"/>'),matrixFill("FFFFFF","","srgbClr"),"<a:noFill/>"]});}
  const fixture=matrixFixture({fill,border,scheme}),parts=readGeneratedZip((await generatePptx(document)).bytes);parts.set("ppt/tableStyles.xml",Buffer.from(fixture.parts["deck/tableStyles.xml"]));parts.set("ppt/theme/theme1.xml",Buffer.from(parts.get("ppt/theme/theme1.xml").toString().replace(/<a:fmtScheme[\s\S]*?<\/a:fmtScheme>/,scheme)));parts.set("ppt/slides/slide1.xml",Buffer.from(parts.get("ppt/slides/slide1.xml").toString().replace(/<a:tblPr[^>]*>[\s\S]*?<\/a:tblPr>/,`<a:tblPr><a:tableStyleId>${TABLE_STYLE_ID}</a:tableStyleId></a:tblPr>`).replace(/<a:tcPr>[\s\S]*?<\/a:tcPr>/g,"<a:tcPr/>")));
  const bytes=zip([...parts].map(([name,body])=>({name,body}))).bytes,file=path.join(directory,`${name}.pptx`);await writeFile(file,bytes);sources.push(file);hashes.push({name,sourceSha256:hash(bytes)});const session=await importPptx(bytes,{projection:"content"});try{expected.push({name,table:normalizePresentation(session.document).slides[0].elements[0],provenance:session.report.tableStyleInheritance});}finally{session.dispose();}
}
const expectedFile=path.join(directory,"expected.json");await writeFile(expectedFile,JSON.stringify(expected));const logs=[];
for(const [i,source]of sources.entries()){
  const name=expected[i].name,profile=`-env:UserInstallation=${pathToFileURL(path.join(directory,`profile-${name}`)).href}`;
  const attempt=async(stage,args,output)=>{try{const result=await execute(executable,[profile,...args],options);await readFile(output);logs.push({name,stage,stdout:result.stdout});return true;}catch(error){failures.push({name,stage,code:error.code??null,message:String(error.stderr??error.message).slice(0,4000)});return false;}};
  const resaved=await attempt("native-resave",["--headless","--convert-to","pptx:Impress Office Open XML","--outdir",nativeDir,source],path.join(nativeDir,`${name}.pptx`));
  if(resaved)await attempt("saved-pdf",["--headless","--convert-to","pdf:impress_pdf_Export","--outdir",pdfDir,path.join(nativeDir,`${name}.pptx`)],path.join(pdfDir,`${name}.pdf`));
  await attempt("source-pdf",["--headless","--convert-to","pdf:impress_pdf_Export","--outdir",controlDir,source],path.join(controlDir,`${name}.pdf`));
}
const independent=JSON.parse((await execute("python3",["scripts/research/oracles/theme-matrix.py",expectedFile,directory],options)).stdout);
for(const h of hashes)for(const[key,file]of [["nativeSha256",path.join(nativeDir,`${h.name}.pptx`)],["pdfSha256",path.join(pdfDir,`${h.name}.pdf`)],["sourceControlPdfSha256",path.join(controlDir,`${h.name}.pdf`)]]){try{h[key]=hash(await readFile(file));}catch(error){if(error.code!=="ENOENT")throw error;}}
const evidence={schemaVersion:1,recordedAt:new Date().toISOString(),application,platform:`${platform()} ${release()} ${arch()}`,artifactsDirectory:path.relative(process.cwd(),directory),scope:"Theme matrix solid/no-fill and table line references: independent source index/color/alpha oracle, native direct XML and source-control/resaved PDF text/fill pixel observations. Other paints/effects, full line-pixel/typography/native interaction and source-style writes remain separate gates.",hashes,independent,logs,failures,interactiveEditing:"untested",sourceStyleWrites:"blocked"};await writeFile(path.join(directory,"evidence.json"),JSON.stringify(evidence,null,2)+"\n");await writeFile("docs/testing/theme-matrix-evidence.json",JSON.stringify(evidence,null,2)+"\n");console.log(JSON.stringify({directory,application,scenes:independent.scenes.map(s=>({name:s.name,native:s.nativeDirectFormatting.status,pdf:s.pdf.map(p=>p.fillPixels)}))}));
