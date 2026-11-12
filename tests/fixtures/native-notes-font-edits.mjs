import { nativeNotesFontBytes } from "./native-notes-fonts.mjs";
import { readGeneratedZip } from "../support/read-generated-zip.mjs";
import { zip } from "../../packages/pptx-parser/test/zip-fixture.mjs";

export const nativeNotesEditTargets = [3,5,6,7,8,10,11,12,13,14,15];
/** Source-only two-run/local-style regression; does not execute native PowerPoint editing. */
export async function nativeNotesFontEditsSourceBytes(variant="A",stage="append"){
  if(!["append","bold","italic","regular"].includes(stage))throw Error("Unknown source notes stage");
  const fixture=await nativeNotesFontBytes(variant),parts=readGeneratedZip(fixture.bytes);
  for(const index of nativeNotesEditTargets){
    const part=`ppt/notesSlides/notesSlide${index}.xml`;let count=0;
    const next=parts.get(part).toString().replace(/<p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes Text"[\s\S]*?<\/p:sp>/,shape=>shape.replace(/<a:r><a:rPr\b([^>]*)>([\s\S]*?)<\/a:rPr><a:t>[\s\S]*?<\/a:t><\/a:r>/,run=>{
      count++;const properties=/<a:rPr\b([^>]*)>([\s\S]*?)<\/a:rPr>/.exec(run);if(!properties)throw Error("Expected notes run properties");
      const flags=stage==="bold"?' b="1"':stage==="italic"?' b="0" i="1"':stage==="regular"?' b="0" i="0"':"";
      const added=`<a:r><a:rPr${properties[1].replace(/lang="[^"]*"/,'lang="en-US"')}${flags}>${properties[2]}</a:rPr><a:t>Ω</a:t></a:r>`;
      return run+added;
    }));
    // Generated empty rPr uses explicit open/close tags in this named source assembler.
    if(count!==1)throw Error("Expected one notes body run");parts.set(part,Buffer.from(next));
  }
  return{...fixture,bytes:zip([...parts].map(([name,body])=>({name,body}))).bytes,stage,targets:nativeNotesEditTargets};
}
