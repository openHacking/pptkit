import assert from "node:assert/strict";
import { createPresentation } from "../../packages/core/dist/index.js";
import { generatePptx } from "../../packages/pptx-exporter/dist/index.js";
import { inspectPptx, inspectXmlPart } from "../../packages/pptx-parser/dist/index.js";
import { readGeneratedZip } from "../../tests/support/read-generated-zip.mjs";
import { zip } from "../../packages/pptx-parser/test/zip-fixture.mjs";
import { nativeDiagramSpecimen } from "./native-diagram-fixture.mjs";

// Owned fresh research package only: never patch an arbitrary source or pretend
// to preserve an imported deck. Component bytes are supplied explicitly offline.
export async function buildNativeDiagramCandidate(candidate, components, options = {}) {
  if (options.drawingId !== undefined && options.drawingId !== 3) throw new TypeError("The fresh specimen owns drawing ID 3.");
  const specimen = nativeDiagramSpecimen(candidate, { ...options, drawingId: 3 });
  const D = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
  const captured = {};
  for (const name of ["layout", "colors", "style"]) {
    if (!(components?.[name] instanceof Uint8Array)) throw new TypeError(`Explicit ${name} XML bytes are required.`);
    if (components[name].byteLength > 2000000) throw new RangeError("Component exceeds the research input budget.");
    captured[name] = new Uint8Array(components[name]);
  }
  for (const [name, local] of [["layout", "layoutDef"], ["colors", "colorsDef"], ["style", "styleDef"]]) {
    const inspected = await inspectXmlPart(captured[name], { partName: name, limits: { maxInputBytes: 2000000 } });
    assert.equal(inspected.root.namespaceUri, D); assert.equal(inspected.root.local, local);
  }
  const document = createPresentation({ id: "native-diagram-admission", metadata: { title: "PPTKit native diagram admission specimen" } });
  document.addSlide({ id: "admission-slide" }).addElement({ id: "admission-control", type: "text", content: "PPTKit native diagram control", box: { x: 40, y: 25, width: 600, height: 35 } });
  const base = await generatePptx(document), parts = readGeneratedZip(base.bytes);
  const slide = parts.get("ppt/slides/slide1.xml").toString();
  assert.equal((slide.match(/<p:cNvPr /g) ?? []).length, 2); assert.match(slide, /<p:cNvPr id="2"/);
  parts.set("ppt/slides/slide1.xml", Buffer.from(slide.replace("</p:spTree>", specimen.frameXml + "</p:spTree>")));
  const mime = "application/vnd.openxmlformats-officedocument.drawingml.";
  const entries = [
    ["data", specimen.dataXml, "diagramData", "diagramData"],
    ["layout", captured.layout, "diagramLayout", "diagramLayout"],
    ["colors", captured.colors, "diagramColors", "diagramColors"],
    ["style", captured.style, "diagramStyle", "diagramQuickStyle"],
  ];
  let types = parts.get("[Content_Types].xml").toString(), relations = parts.get("ppt/slides/_rels/slide1.xml.rels").toString();
  for (const [name, bytes, content, role] of entries) {
    const file = `${name}1.xml`, identifier = `pptkitDiagram${name[0].toUpperCase()}${name.slice(1)}`;
    parts.set(`ppt/diagrams/${file}`, Buffer.from(bytes));
    types = types.replace("</Types>", `<Override PartName="/ppt/diagrams/${file}" ContentType="${mime}${content}+xml"/></Types>`);
    relations = relations.replace("</Relationships>", `<Relationship Id="${identifier}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${role}" Target="../diagrams/${file}"/></Relationships>`);
  }
  if (specimen.cacheXml !== undefined) {
    parts.set("ppt/diagrams/drawing1.xml", Buffer.from(specimen.cacheXml));
    types = types.replace("</Types>", '<Override PartName="/ppt/diagrams/drawing1.xml" ContentType="application/vnd.ms-office.drawingml.diagramDrawing+xml"/></Types>');
    relations = relations.replace("</Relationships>", '<Relationship Id="pptkitDiagramDrawing" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="../diagrams/drawing1.xml"/></Relationships>');
  }
  parts.set("[Content_Types].xml", Buffer.from(types)); parts.set("ppt/slides/_rels/slide1.xml.rels", Buffer.from(relations));
  const bytes = zip([...parts].map(([name, body]) => ({ name, body }))).bytes, inspection = await inspectPptx(bytes);
  assert.equal(inspection.slides[0].diagrams.length, 1);
  assert.equal(inspection.slides[0].diagrams[0].components.find(c => c.role === "layout").layoutIdentifier, specimen.layoutIdentifier);
  assert.ok(inspection.slides[0].diagrams[0].components.every(c => c.status === "located"));
  assert.equal(inspection.diagramInventory.dataModels[0].status, "inspected");
  assert.ok(inspection.diagramInventory.parts.every(part => part.imageRelationships.length === 0), "Resource-dependent components are outside this specimen profile.");
  return { bytes: new Uint8Array(bytes), inspection, specimen };
}
