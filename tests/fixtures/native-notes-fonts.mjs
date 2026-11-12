import { createPresentation } from "../../packages/core/dist/index.js";
import { generatePptx } from "../../packages/pptx-exporter/dist/index.js";
import { readGeneratedZip } from "../support/read-generated-zip.mjs";
import { zip } from "../../packages/pptx-parser/test/zip-fixture.mjs";
import { nativeFontFamilies as F } from "../../scripts/research/font-native-cases.mjs";

const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const xml = value => value.replaceAll("&", "&amp;").replaceAll('"', "&quot;").replaceAll("<", "&lt;");
const slots = values => ["latin", "ea", "cs"].map((slot, i) => values[i] === null ? "" : `<a:${slot} typeface="${xml(values[i])}"/>`).join("");
const transform = (x, y, width, height) => `<a:xfrm><a:off x="${x*12700}" y="${y*12700}"/><a:ext cx="${width*12700}" cy="${height*12700}"/></a:xfrm>`;
const shape = (id, name, placeholder, geometry, text, list = "") => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr>${placeholder}</p:nvPr></p:nvSpPr><p:spPr>${geometry}</p:spPr><p:txBody><a:bodyPr/><a:lstStyle>${list}</a:lstStyle>${text}</p:txBody></p:sp>`;
const p = text => `<a:p><a:r><a:rPr lang="en-US" sz="1600">${slots([F.latin,F.latin,F.latin])}</a:rPr><a:t>${xml(text)}</a:t></a:r></a:p>`;
const samples = [["latin","Abc","en-US","Latn"],["sc","骨直令返","zh-CN","Hans"],["tc","骨直令返","zh-TW","Hant"],["ja","骨直令返","ja","Jpan"],["ko","한글","ko","Hang"],["ko-kr","한글","ko-KR","Hang"],["ar","السَّلَامُ","ar","Arab"],["he","שָׁלוֹם","he","Hebr"],["greek","Ω","el","Grek"]];
export const notesFontMaps = {
  A: { Latn: F.latin, Hans: F.sc, Hant: F.tc, Jpan: F.ja, Hang: F.ko, Arab: F.ar, Hebr: F.he, Grek: F.greek },
  B: { Latn: F.greek, Hans: F.tc, Hant: F.sc, Jpan: F.sc, Hang: F.generic, Arab: F.generic, Hebr: F.generic, Grek: F.latin },
};

/** Native research source assembler: author base then patch independent notes theme/master/slots; not fresh-author round trip. */
export async function nativeNotesFontBytes(variant = "A") {
  if (!Object.hasOwn(notesFontMaps, variant)) throw Error("Unknown named notes-font variant");
  const map = notesFontMaps[variant], controls = samples.map(([id,text,language,script]) => ({ id, text, language, script, localSlots: [null,null,null] }));
  for (const id of ["tc","ar","he"]) { const c = controls.find(c => c.id === id); controls.push({ ...c, id: `local-${id}`, localSlots: [F.latin,F.generic,F.generic] }); }
  controls.push({ ...controls[2], id: "empty-ea-tc", localSlots: [null,"",null] });
  controls.push({ ...controls[2], id: "bad-ea-tc", localSlots: [null,"+mn-unknown",null] });
  controls.push({ ...controls[2], id: "major-tc", localSlots: ["+mj-lt","+mj-ea","+mj-cs"] });
  const document = createPresentation({ size: { width:720,height:260 }, theme:{fonts:{body:F.latin,heading:F.latin}} });
  for (const c of controls) {
    const slide=document.addSlide({id:`notes-${variant}-${c.id}`, notes:[{runs:[{text:c.text,style:{fontFamily:F.latin,fontSize:18,language:c.language}}]}]});
    slide.addElement({type:"text",name:"SlideLabel",box:{x:20,y:20,width:680,height:40},content:[{runs:[{text:`notes-${variant}-${c.id}`,style:{fontFamily:F.latin,fontSize:16}}]}]});
    slide.addElement({type:"text",name:"SlideProbe",box:{x:20,y:90,width:680,height:110},content:[{runs:[{text:c.text,style:{fontFamily:F.generic,fontSize:48,language:c.language}}]}]});
  }
  const authorBase=(await generatePptx(document)).bytes, parts=readGeneratedZip(authorBase), themePart="ppt/theme/independent-notes-theme.xml";
  let theme=parts.get("ppt/theme/theme2.xml").toString().replace('name="PPTKit"',`name="PPTKit Notes ${variant}"`);
  for (const kind of ["major","minor"]) theme=theme.replace(new RegExp(`<a:${kind}Font>[\\s\\S]*?</a:${kind}Font>`),`<a:${kind}Font>${slots([map.Latn,"",""])}${Object.entries(map).map(([script,family])=>`<a:font script="${script}" typeface="${xml(family)}"/>`).join("")}</a:${kind}Font>`);
  parts.set(themePart,Buffer.from(theme));
  // Keep theme2 deliberately unused: the actual relationship must be followed.
  const relPart="ppt/notesMasters/_rels/notesMaster1.xml.rels";
  parts.set(relPart,Buffer.from(parts.get(relPart).toString().replace('Target="../theme/theme2.xml"','Target="../theme/independent-notes-theme.xml"')));
  parts.set("[Content_Types].xml",Buffer.from(parts.get("[Content_Types].xml").toString().replace('</Types>',`<Override PartName="/${themePart}" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>`)));
  let master=parts.get("ppt/notesMasters/notesMaster1.xml").toString();
  // This source research control owns its independent master geometry. Keep it
  // separate from fresh Exporter defaults, including future default placeholders.
  master=master.replace(/<p:sp>[\s\S]*?<\/p:sp>/g, "");
  const masterBody=shape(5,"Notes Body Master",'<p:ph type="body" idx="99"/>',transform(36,360,468,280),p("MASTER PROMPT NEVER INHERIT"),'<a:lvl1pPr><a:defRPr sz="4000"/></a:lvl1pPr>');
  // PowerPoint prints no slide image when this master placeholder has only xfrm.
  // Keep the local placeholder inherited; the visible image needs master geometry.
  const image=shape(2,"Slide Image",'<p:ph type="sldImg" idx="2"/>',transform(36,36,468,169)+'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>','<a:p><a:endParaRPr lang="en-US"/></a:p>');
  const label=shape(3,"Notes Label","",transform(36,310,468,36),p(`Independent Notes ${variant}`));
  master=master.replace('</p:spTree>',image+label+masterBody+'</p:spTree>').replace(/<p:notesStyle>[\s\S]*?<\/p:notesStyle>/,`<p:notesStyle><a:lvl1pPr><a:defRPr lang="en-US" sz="1800">${slots(["+mn-lt","+mn-ea","+mn-cs"])}<a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:defRPr></a:lvl1pPr></p:notesStyle>`);
  parts.set("ppt/notesMasters/notesMaster1.xml",Buffer.from(master));
  for (const [i,c] of controls.entries()) {
    const file=`ppt/notesSlides/notesSlide${i+1}.xml`;let note=parts.get(file).toString();
    note=note.replace(/<p:sp>[\s\S]*?<\/p:sp>/g, shape => shape.includes('<p:ph type="sldImg"') ? "" : shape);
    note=note.replace('type="body" idx="1"','type="body" idx="7"').replace(/<a:rPr\b[^>]*>[\s\S]*?<\/a:rPr>/,`<a:rPr lang="${c.language}">${slots(c.localSlots)}</a:rPr>`);
    const localLabel=shape(3,"Notes Control Label","",transform(36,270,468,36),p(`notes-${variant}-${c.id}`));
    const localImage=shape(4,"Slide Image Local",'<p:ph type="sldImg" idx="2"/>',"",'<a:p><a:endParaRPr lang="en-US"/></a:p>');
    note=note.replace('</p:spTree>',localLabel+localImage+'</p:spTree>');
    parts.set(file,Buffer.from(note));
  }
  return {authorBase,bytes:zip([...parts].map(([name,body])=>({name,body}))).bytes,variant,controls,map,themePart,sourceAssembly:"Generated author base; patched notes master/body geometry including slide-image rect, type-vs-index, independent font theme/relationship and local slots; not fresh-author round trip."};
}
