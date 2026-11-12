import { body, contentFixture } from "./content-fixture.mjs";
import { encodeStructure } from "./structure-fixture.mjs";
import { relationships } from "./opc-fixture.mjs";

export const notesShape = ({ id = "2", role = "body", text = body(), placeholder = true, props = "", name = "Any name" } = {}) =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr>${placeholder ? `<p:ph type="${role}" idx="1"/>` : ""}</p:nvPr></p:nvSpPr><p:spPr>${props}</p:spPr>${text}</p:sp>`;
export const notesTree = content => `<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name="Root"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${content}</p:spTree></p:cSld>`;
export function notesFixture(content = notesShape(), strict = false) {
  const f = contentFixture(undefined, strict);
  f.parts["deck/notes/z.xml"] = f.p("notes", notesTree(content), `xmlns:a="${f.A}"`);
  f.types["deck/notes/z.xml"] = "presentationml.notesSlide";
  f.parts["deck/_rels/z.xml.rels"] = relationships([f.rel("layout", "slideLayout", "layout.xml"), f.rel("notes", "notesSlide", "notes/z.xml")]);
  f.parts["deck/notes/_rels/z.xml.rels"] = relationships([f.rel("back", "slide", "../z.xml"), f.rel("master", "notesMaster", "../notes-master.xml")]);
  f.parts["deck/notes-master.xml"] = f.p("notesMaster", `${notesTree(notesShape({ id: "5", text: body("<a:p><a:r><a:t>Master must not become notes</a:t></a:r></a:p>") }))}<p:notesStyle><a:lvl1pPr><a:defRPr sz="9900"/></a:lvl1pPr></p:notesStyle>`, `xmlns:a="${f.A}"`);
  f.types["deck/notes-master.xml"] = "presentationml.notesMaster";
  f.parts["deck/_rels/notes-master.xml.rels"] = relationships([f.rel("theme", "theme", "theme.xml")]);
  f.parts["deck/main.data"] = f.parts["deck/main.data"].replace("<p:sldIdLst>", '<p:notesMasterIdLst><p:notesMasterId r:id="notesMaster"/></p:notesMasterIdLst><p:sldIdLst>');
  f.parts["deck/_rels/main.data.rels"] = relationships([f.rel("a", "slide", "a.xml"), f.rel("m", "slideMaster", "master.xml"), f.rel("z", "slide", "z.xml"), f.rel("notesMaster", "notesMaster", "notes-master.xml")]);
  return f;
}
export const notesBytes = (content, strict) => encodeStructure(notesFixture(content, strict));
