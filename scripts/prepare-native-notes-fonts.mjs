import assert from "node:assert/strict";
import { mkdir, mkdtemp, readFile, writeFile } from "node:fs/promises";
import { createHash } from "node:crypto";
import path from "node:path";
import { normalizePresentation, resolveTextFontSegments } from "../packages/core/dist/index.js";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { nativeNotesFontBytes } from "../tests/fixtures/native-notes-fonts.mjs";

const hash = bytes => createHash("sha256").update(bytes).digest("hex"), options = { projection:"content", notesTextProfile:"master-style-v1",textFontProfile:"script-families-v3",editSaveProfile:"text-runs-v3" };
const fonts=JSON.parse(await readFile("docs/testing/real-font-native-evidence.json","utf8")).fonts;
for (const font of fonts) assert.equal(hash(await readFile(font.file)),font.sha256);
await mkdir(".tmp/native-notes-fonts",{recursive:true});const directory=await mkdtemp(path.resolve(".tmp/native-notes-fonts/run-"));
for (const variant of ["A","B"]) {
  const fixture=await nativeNotesFontBytes(variant), folder=path.join(directory,variant);await mkdir(folder);
  const session=await importPptx(fixture.bytes,options);
  try {
    const snapshot=normalizePresentation(session.document), projections=snapshot.slides.map((slide,i)=>{
      const run=slide.notes[0].runs[0];assert.equal(run.text,fixture.controls[i].text);assert.equal(run.style.fontSize,40);assert.equal(run.style.language,fixture.controls[i].language);
      return {id:fixture.controls[i].id,notes:slide.notes,selected:resolveTextFontSegments(run.text,run.style,snapshot.theme)};
    });
    const plan=await session.prepareSave({mode:"no-op"});assert.equal(plan.status,"ready");assert.deepEqual(Buffer.from((await session.save(plan.plan)).bytes),Buffer.from(fixture.bytes));
    await writeFile(path.join(folder,"author-base.pptx"),fixture.authorBase);await writeFile(path.join(folder,"notes-fonts.pptx"),fixture.bytes);
    await writeFile(path.join(folder,"expected.json"),JSON.stringify({variant,sourceAssembly:fixture.sourceAssembly,sourceSha256:hash(fixture.bytes),authorBaseSha256:hash(fixture.authorBase),controls:fixture.controls,map:fixture.map,themePart:fixture.themePart,fonts,options,projections,properties:session.report.textStyleInheritance.properties},null,2)+"\n");
  }finally{session.dispose();}
}
await writeFile(".tmp/native-notes-fonts/latest.txt",directory);console.log(JSON.stringify({directory,groups:2,controlsPerGroup:15}));
