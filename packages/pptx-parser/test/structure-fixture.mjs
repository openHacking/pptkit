import { pack, contentTypes, relationships } from "./opc-fixture.mjs";
export function structureFixture(strict = false) {
  const P = strict ? "http://purl.oclc.org/ooxml/presentationml/main" : "http://schemas.openxmlformats.org/presentationml/2006/main";
  const R = strict ? "http://purl.oclc.org/ooxml/officeDocument/relationships" : "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const A = strict ? "http://purl.oclc.org/ooxml/drawingml/main" : "http://schemas.openxmlformats.org/drawingml/2006/main";
  const p = (name, content = "", attrs = "") => `<p:${name} xmlns:p="${P}" xmlns:r="${R}" ${attrs}>${content}</p:${name}>`;
  const rel = (Id, type, Target) => ({ Id, Type: `${R}/${type}`, Target });
  const parts = {
    "deck/main.data": p("presentation", '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="m"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id="300" r:id="z"/><p:sldId id="256" r:id="a"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/>'),
    "_rels/.rels": relationships([rel("office", "officeDocument", "deck/main.data")]),
    "deck/_rels/main.data.rels": relationships([rel("a", "slide", "a.xml"), rel("m", "slideMaster", "master.xml"), rel("z", "slide", "z.xml")]),
    "deck/a.xml": p("sld", '<p:cSld name="Second"/>'),
    "deck/z.xml": p("sld", '<p:cSld name="First"/>', 'show="0"'),
    "deck/_rels/a.xml.rels": relationships([rel("layout", "slideLayout", "layout.xml")]),
    "deck/_rels/z.xml.rels": relationships([rel("layout", "slideLayout", "layout.xml")]),
    "deck/layout.xml": p("sldLayout", '<p:cSld name="Shared layout"/>'),
    "deck/_rels/layout.xml.rels": relationships([rel("master", "slideMaster", "master.xml")]),
    "deck/master.xml": p("sldMaster", '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="layout"/></p:sldLayoutIdLst>'),
    "deck/_rels/master.xml.rels": relationships([rel("layout", "slideLayout", "layout.xml"), rel("theme", "theme", "theme.xml")]),
    "deck/theme.xml": `<a:theme xmlns:a="${A}"/>`,
  };
  const types = { "deck/main.data": "presentationml.presentation.main", "deck/a.xml": "presentationml.slide", "deck/z.xml": "presentationml.slide", "deck/layout.xml": "presentationml.slideLayout", "deck/master.xml": "presentationml.slideMaster", "deck/theme.xml": "theme" };
  return { parts, types, P, R, A, p, rel };
}
export function encodeStructure(f) {
  return pack(f.parts, contentTypes(undefined, Object.entries(f.types).map(([PartName, type]) => ({ PartName: `/${PartName}`, ContentType: `application/vnd.openxmlformats-officedocument.${type}+xml` }))));
}
