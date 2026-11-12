import { structureFixture, encodeStructure } from "./structure-fixture.mjs";
import { relationships } from "./opc-fixture.mjs";

export function smartArtFixture(strict = false) {
  const f = structureFixture(strict), D = f.A.replace(/\/main$/, "/diagram");
  const frame = (id = "2", refs = 'r:dm="dm" r:lo="lo" r:cs="cs" r:qs="qs"', uri = D) => `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="Diagram ${id}"/></p:nvGraphicFramePr><a:graphic><a:graphicData uri="${uri}"><d:relIds ${refs}/></a:graphicData></a:graphic></p:graphicFrame>`;
  const setTree = (body, attrs = "") => { f.parts["deck/z.xml"] = f.p("sld", `<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1"/></p:nvGrpSpPr>${body}</p:spTree></p:cSld>`, `xmlns:a="${f.A}" xmlns:d="${D}" ${attrs}`); };
  const relations = [f.rel("layout", "slideLayout", "layout.xml"), ...[["dm", "diagramData", "data"], ["lo", "diagramLayout", "layout"], ["cs", "diagramColors", "colors"], ["qs", "diagramQuickStyle", "style"]].map(([id, role, name]) => f.rel(id, role, `diagrams/${name}.xml`))];
  const setRelations = () => { f.parts["deck/_rels/z.xml.rels"] = relationships(relations); };
  for (const [name, root, type] of [["data", "dataModel", "diagramData"], ["layout", "layoutDef", "diagramLayout"], ["colors", "colorsDef", "diagramColors"], ["style", "styleDef", "diagramStyle"]]) {
    f.parts[`deck/diagrams/${name}.xml`] = `<d:${root} xmlns:d="${D}"${name === "layout" ? ' uniqueId="urn:pptkit:test:unknown-layout"' : ""}/>`;
    f.types[`deck/diagrams/${name}.xml`] = `drawingml.${type}`;
  }
  setTree(frame()); setRelations();
  return Object.assign(f, { D, frame, setTree, relations, setRelations, bytes: () => encodeStructure(f) });
}
