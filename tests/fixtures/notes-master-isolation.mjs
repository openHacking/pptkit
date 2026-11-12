import { nativeNotesFontBytes, notesFontMaps } from "./native-notes-fonts.mjs";
import { readGeneratedZip } from "../support/read-generated-zip.mjs";
import { zip } from "../../packages/pptx-parser/test/zip-fixture.mjs";

const relType="http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
const override=(part,type)=>`<Override PartName="/${part}" ContentType="application/vnd.openxmlformats-officedocument.${type}+xml"/>`;
const replaceOnce=(source,from,to)=>{if(source.split(from).length!==2)throw Error(`Expected one source occurrence: ${from}`);return source.replace(from,to);};

/** NON-CONFORMING source control: two notes masters, interleaved order. Listed duplicate must reject; unlisted actual edge must be diagnosed. */
export async function notesMasterIsolationSourceBytes({ listedSecond=false }={}){
  const a=await nativeNotesFontBytes("A"),b=await nativeNotesFontBytes("B"),parts=readGeneratedZip(a.bytes),second=readGeneratedZip(b.bytes),controls=[];
  const masterB="ppt/notesMasters/notesMaster2.xml",themeB="ppt/theme/independent-notes-theme-B.xml";
  parts.set(masterB,second.get("ppt/notesMasters/notesMaster1.xml"));parts.set(themeB,second.get(b.themePart));
  parts.set("ppt/notesMasters/_rels/notesMaster2.xml.rels",Buffer.from(replaceOnce(second.get("ppt/notesMasters/_rels/notesMaster1.xml.rels").toString(),'Target="../theme/independent-notes-theme.xml"','Target="../theme/independent-notes-theme-B.xml"')));
  parts.set("docProps/app.xml",Buffer.from(replaceOnce(parts.get("docProps/app.xml").toString(),"<Slides>15</Slides>","<Slides>30</Slides>")));
  let additions=override(masterB,"presentationml.notesMaster")+override(themeB,"theme"),relationships="",ids="";
  for(let i=1;i<=15;i++){
    const target=i+15;
    parts.set(`ppt/slides/slide${target}.xml`,second.get(`ppt/slides/slide${i}.xml`));
    parts.set(`ppt/slides/_rels/slide${target}.xml.rels`,Buffer.from(replaceOnce(second.get(`ppt/slides/_rels/slide${i}.xml.rels`).toString(),`Target="../notesSlides/notesSlide${i}.xml"`,`Target="../notesSlides/notesSlide${target}.xml"`)));
    parts.set(`ppt/notesSlides/notesSlide${target}.xml`,second.get(`ppt/notesSlides/notesSlide${i}.xml`));
    let notesRel=replaceOnce(second.get(`ppt/notesSlides/_rels/notesSlide${i}.xml.rels`).toString(),`Target="../slides/slide${i}.xml"`,`Target="../slides/slide${target}.xml"`);
    notesRel=replaceOnce(notesRel,'Target="../notesMasters/notesMaster1.xml"','Target="../notesMasters/notesMaster2.xml"');parts.set(`ppt/notesSlides/_rels/notesSlide${target}.xml.rels`,Buffer.from(notesRel));
    additions+=override(`ppt/slides/slide${target}.xml`,"presentationml.slide")+override(`ppt/notesSlides/notesSlide${target}.xml`,"presentationml.notesSlide");
    relationships+=`<Relationship Id="rIdIsolatedSlide${i}" Type="${relType}slide" Target="slides/slide${target}.xml"/>`;
    ids+=`<p:sldId id="${255+i}" r:id="rId${i+1}"/><p:sldId id="${270+i}" r:id="rIdIsolatedSlide${i}"/>`;
    for(const variant of ["A","B"]){const physical=variant==="A"?i:target;controls.push({...a.controls[i-1],variant,sourceSlide:physical,notesPart:`ppt/notesSlides/notesSlide${physical}.xml`,masterPart:variant==="A"?"ppt/notesMasters/notesMaster1.xml":masterB,themePart:variant==="A"?a.themePart:themeB,map:notesFontMaps[variant]});}
  }
  let presentation=parts.get("ppt/presentation.xml").toString();if(listedSecond)presentation=replaceOnce(presentation,'</p:notesMasterIdLst>','<p:notesMasterId r:id="rIdIsolatedNotesMaster"/></p:notesMasterIdLst>');
  let replaced=0;presentation=presentation.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/,()=>{replaced++;return`<p:sldIdLst>${ids}</p:sldIdLst>`;});if(replaced!==1)throw Error("Expected one slide order list");parts.set("ppt/presentation.xml",Buffer.from(presentation));
  const relPart="ppt/_rels/presentation.xml.rels";parts.set(relPart,Buffer.from(replaceOnce(parts.get(relPart).toString(),'</Relationships>',relationships+`<Relationship Id="rIdIsolatedNotesMaster" Type="${relType}notesMaster" Target="notesMasters/notesMaster2.xml"/></Relationships>`)));
  parts.set("[Content_Types].xml",Buffer.from(replaceOnce(parts.get("[Content_Types].xml").toString(),'</Types>',additions+'</Types>')));
  return{nonConforming:true,listedSecond,bytes:zip([...parts].map(([name,body])=>({name,body}))).bytes,controls,sourceAssembly:"NON-CONFORMING multi-notes-master adversarial source. Two generated author bases patched/merged, actual A/B edges, interleaved order and retained unused theme2; diagnostic source transport only, never standard/native admission."};
}

/** One actual notes master, with B's supported per-notes fontScheme overrides. Cardinality control, not full schema/native admission. */
export async function notesOverrideIsolationSourceBytes({ distinctParts=false }={}){
  const f=await notesMasterIsolationSourceBytes(),parts=readGeneratedZip(f.bytes),overridePart="ppt/theme/notes-font-override-B.xml";
  const bTheme=parts.get("ppt/theme/independent-notes-theme-B.xml").toString(),scheme=/<a:fontScheme\b[^>]*>[\s\S]*?<\/a:fontScheme>/.exec(bTheme)?.[0];if(!scheme)throw Error("Expected B font scheme");
  parts.set(overridePart,Buffer.from(`<a:themeOverride xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${scheme}</a:themeOverride>`));
  for(let i=16;i<=30;i++){
    const name=`ppt/notesSlides/_rels/notesSlide${i}.xml.rels`;let rels=replaceOnce(parts.get(name).toString(),'Target="../notesMasters/notesMaster2.xml"','Target="../notesMasters/notesMaster1.xml"');
    rels=replaceOnce(rels,'</Relationships>',`<Relationship Id="rIdIsolatedFontOverride" Type="${relType}themeOverride" Target="../theme/notes-font-override-B.xml"/></Relationships>`);parts.set(name,Buffer.from(rels));
  }
  for(const name of ["ppt/notesMasters/notesMaster2.xml","ppt/notesMasters/_rels/notesMaster2.xml.rels","ppt/theme/independent-notes-theme-B.xml"])parts.delete(name);
  const relName="ppt/_rels/presentation.xml.rels",input=parts.get(relName).toString(),removed=input.replace(/<Relationship\b[^>]*Id="rIdIsolatedNotesMaster"[^>]*\/>/,"");if(input===removed)throw Error("Expected second presentation master edge");parts.set(relName,Buffer.from(removed));
  let types=parts.get("[Content_Types].xml").toString();for(const part of ["ppt/notesMasters/notesMaster2.xml","ppt/theme/independent-notes-theme-B.xml"]){const pattern=new RegExp(`<Override PartName="/${part.replaceAll(".","\\.")}"[^>]*/>`);if(!pattern.test(types))throw Error("Expected removed content type");types=types.replace(pattern,"");}parts.set("[Content_Types].xml",Buffer.from(replaceOnce(types,'</Types>',override(overridePart,"themeOverride")+'</Types>')));
  const output=distinctParts?distinctNotesFontOverrideParts(parts):parts;
  return{...f,nonConforming:false,distinctParts,bytes:zip([...output].map(([name,body])=>({name,body}))).bytes,controls:f.controls.map(c=>{const actualOverride=distinctParts?`ppt/theme/notes-font-override-${c.sourceSlide}.xml`:overridePart;return{...c,masterPart:"ppt/notesMasters/notesMaster1.xml",masterThemePart:"ppt/theme/independent-notes-theme.xml",...(c.variant==="B"?{themePart:actualOverride,overridePart:actualOverride}:{} )};}),sourceAssembly:`Generated source controls merged into one package with a single listed/actual notes master, interleaved A/B notes and B fontScheme overrides (${distinctParts?"one independent override part per B notes":"one shared override part"}). Notes-master cardinality is conformant; full schema/native output remains unverified.`};
}

/** Research-only source conversion. Native PowerPoint admission is verified separately. */
export function distinctNotesFontOverrideParts(input){
  const parts=new Map(input),shared="ppt/theme/notes-font-override-B.xml",body=parts.get(shared);
  if(!body)throw Error("Expected shared notes font override");
  let types=parts.get("[Content_Types].xml").toString();
  const sharedType=new RegExp(`<Override PartName="/${shared.replaceAll(".","\\.")}"[^>]*/>`);
  if(!sharedType.test(types))throw Error("Expected shared override content type");
  types=types.replace(sharedType,"");parts.delete(shared);
  for(let index=16;index<=30;index++){
    const target=`ppt/theme/notes-font-override-${index}.xml`,relation=`ppt/notesSlides/_rels/notesSlide${index}.xml.rels`;
    if(parts.has(target))throw Error("Distinct override part already exists");
    parts.set(target,body);
    parts.set(relation,Buffer.from(replaceOnce(parts.get(relation).toString(),'Target="../theme/notes-font-override-B.xml"',`Target="../theme/notes-font-override-${index}.xml"`)));
    types=replaceOnce(types,'</Types>',override(target,"themeOverride")+'</Types>');
  }
  parts.set("[Content_Types].xml",Buffer.from(types));return parts;
}
