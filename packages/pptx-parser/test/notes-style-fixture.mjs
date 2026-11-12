import { notesFixture, notesShape, notesTree } from "./notes-fixture.mjs";
import { body } from "./content-fixture.mjs";
import { relationships } from "./opc-fixture.mjs";

export const notesFonts = (latin = "+mn-lt", ea = "+mn-ea", cs = "+mn-cs") => `<a:latin typeface="${latin}"/><a:ea typeface="${ea}"/><a:cs typeface="${cs}"/>`;
export const notesGlobalStyle = `<a:defPPr><a:defRPr lang="zh-TW" sz="1400">${notesFonts()}<a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:defRPr></a:defPPr><a:lvl1pPr><a:defRPr sz="1600"/></a:lvl1pPr>`;
export function notesStyleFixture({ strict = false, global = notesGlobalStyle, placeholder = '<a:lvl1pPr><a:defRPr sz="2200" i="1"/></a:lvl1pPr>', local = "", paragraph = "", run = "", extraMaster = "" } = {}) {
  const text = body(`<a:p>${paragraph ? `<a:pPr>${paragraph}</a:pPr>` : ""}<a:r>${run ? `<a:rPr>${run}</a:rPr>` : ""}<a:t>A中عא</a:t></a:r></a:p>`).replace('<a:lstStyle/>', `<a:lstStyle>${local}</a:lstStyle>`);
  const f = notesFixture(notesShape({ text }), strict);
  const masterBody = notesShape({ id: "5", text: body('<a:p><a:r><a:t>Never inherit MASTER PROMPT</a:t></a:r></a:p>').replace('<a:lstStyle/>', `<a:lstStyle>${placeholder}</a:lstStyle>`) }).replace('idx="1"', 'idx="99"');
  const colorMap = '<p:clrMap bg1="lt1" tx1="accent2" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>';
  f.parts["deck/notes-master.xml"] = f.p("notesMaster", `${notesTree(masterBody + extraMaster)}${colorMap}<p:notesStyle>${global}</p:notesStyle>`, `xmlns:a="${f.A}"`);
  f.parts["deck/_rels/notes-master.xml.rels"] = relationships([f.rel("theme", "theme", "notes-theme.xml")]);
  f.types["deck/notes-theme.xml"] = "theme";
  f.parts["deck/notes-theme.xml"] = `<a:theme xmlns:a="${f.A}" name="Independent Notes"><a:themeElements><a:clrScheme name="Notes"><a:dk1><a:srgbClr val="010203"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:accent1><a:srgbClr val="112233"/></a:accent1><a:accent2><a:srgbClr val="204060"/></a:accent2></a:clrScheme><a:fontScheme name="Notes"><a:majorFont>${notesFonts("Notes Heading", "Notes Heading East", "Notes Heading Complex")}</a:majorFont><a:minorFont>${notesFonts("Notes Latin", "Notes East", "Notes Complex")}<a:font script="Hant" typeface="Notes Traditional"/><a:font script="Arab" typeface="Notes Arabic"/><a:font script="Hebr" typeface="Notes Hebrew"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>`;
  return f;
}
/** differentMaster=true deliberately adds an unlisted second notes master: a non-conforming diagnostic source, not standard authoring admission. */
export function addSecondNotes(f, differentMaster = false) {
  f.parts["deck/_rels/a.xml.rels"] = relationships([f.rel("layout", "slideLayout", "layout.xml"), f.rel("notes", "notesSlide", "notes/a.xml")]);
  f.parts["deck/notes/a.xml"] = f.parts["deck/notes/z.xml"]; f.types["deck/notes/a.xml"] = "presentationml.notesSlide";
  f.parts["deck/notes/_rels/a.xml.rels"] = relationships([f.rel("back", "slide", "../a.xml"), f.rel("master", "notesMaster", differentMaster ? "../notes-master2.xml" : "../notes-master.xml")]);
  if (differentMaster) {
    f.parts["deck/notes-master2.xml"] = f.parts["deck/notes-master.xml"]; f.types["deck/notes-master2.xml"] = "presentationml.notesMaster";
    f.parts["deck/_rels/notes-master2.xml.rels"] = relationships([f.rel("theme", "theme", "notes-theme2.xml")]);
    f.parts["deck/notes-theme2.xml"] = f.parts["deck/notes-theme.xml"].replaceAll("Notes Latin", "Second Notes Latin").replaceAll("Notes Traditional", "Second Notes Traditional"); f.types["deck/notes-theme2.xml"] = "theme";
  }
  return f;
}
