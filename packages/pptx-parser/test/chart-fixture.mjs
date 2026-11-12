import { contentFixture } from "./content-fixture.mjs";
import { contentTypes, relationships, pack } from "./opc-fixture.mjs";
const escape = value => String(value).replaceAll("&", "&amp;").replaceAll("<", "&lt;");
export const CHART = "http://schemas.openxmlformats.org/drawingml/2006/chart";
export function cache(values, { numeric = false, formula, omit = false, count = values.length, indices = values.map((_, i) => i) } = {}) {
  const kind = numeric ? "num" : "str";
  const data = `<c:ptCount val="${count}"/>${values.map((v, i) => `<c:pt idx="${indices[i]}"><c:v>${escape(v)}</c:v></c:pt>`).join("")}`;
  return formula === undefined ? `<c:${kind}Lit>${data}</c:${kind}Lit>` : `<c:${kind}Ref><c:f>${escape(formula)}</c:f>${omit ? "" : `<c:${kind}Cache>${data}</c:${kind}Cache>`}</c:${kind}Ref>`;
}
export function series({ idx = 7, order = 0, name = "Revenue α😀", categories = ["Q1", "Q2"], values = [0, -2.5], refs = false, omit = false, extras = "" } = {}) {
  return `<c:ser><c:idx val="${idx}"/><c:order val="${order}"/><c:tx>${cache([name], refs ? { formula: "Sheet1!$B$1", omit } : {})}</c:tx>${extras}<c:cat>${cache(categories, refs ? { formula: "Sheet1!$A$2:$A$3", omit } : {})}</c:cat><c:val>${cache(values, { numeric: true, ...(refs ? { formula: "Sheet1!$B$2:$B$3", omit } : {}) })}</c:val></c:ser>`;
}
export function frame({ id = 2, ref = "chart", uri = CHART, transform = 'rot="900000" flipH="1"' } = {}) {
  return `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="Chart ${id}" descr="Chart description"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm ${transform}><a:off x="127000" y="254000"/><a:ext cx="5080000" cy="3175000"/></p:xfrm><a:graphic><a:graphicData uri="${uri}"><c:chart xmlns:c="${uri}" r:id="${ref}"/></a:graphicData></a:graphic></p:graphicFrame>`;
}
export const axes = () => `<c:catAx><c:axId val="11"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:crossAx val="12"/></c:catAx><c:valAx><c:axId val="12"/><c:scaling><c:min val="-5"/><c:max val="5"/></c:scaling><c:majorUnit val="2.5"/><c:crossAx val="11"/></c:valAx>`;
export function chartXml({ family = "bar", data = series({ values: family === "pie" ? [0, 2.5] : [0, -2.5] }), before = "", after = "", plotExtra = "", familyExtra = "", external = false } = {}) {
  const bar = family === "bar" ? '<c:barDir val="bar"/><c:grouping val="clustered"/>' : family === "line" ? '<c:grouping val="standard"/>' : "";
  const references = family === "pie" ? "" : '<c:axId val="11"/><c:axId val="12"/>';
  return `<c:chartSpace xmlns:c="${CHART}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:chart>${before}<c:plotArea><c:${family}Chart>${bar}${data}${familyExtra}${references}</c:${family}Chart>${family === "pie" ? "" : axes()}${plotExtra}</c:plotArea>${after}</c:chart>${external ? '<c:externalData r:id="book"><c:autoUpdate val="0"/></c:externalData>' : ""}</c:chartSpace>`;
}
export function workbook({ values = [0, -2.5], categories = ["Q1", "Q2"], name = "Revenue α😀", formula = false, shared = false, error = false } = {}) {
  const S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const str = (ref, value, i) => shared ? `<c r="${ref}" t="s"><v>${i}</v></c>` : `<c r="${ref}" t="inlineStr"><is><t>${escape(value)}</t></is></c>`;
  const parts = {
    "_rels/.rels": relationships([{ Id: "main", Type: `${R}/officeDocument`, Target: "xl/workbook.xml" }]),
    "xl/workbook.xml": `<workbook xmlns="${S}" xmlns:r="${R}"><sheets><sheet name="Sheet1" sheetId="1" r:id="sheet"/></sheets></workbook>`,
    "xl/_rels/workbook.xml.rels": relationships([{ Id: "sheet", Type: `${R}/worksheet`, Target: "sheets/one.xml" }, ...(shared ? [{ Id: "strings", Type: `${R}/sharedStrings`, Target: "strings.xml" }] : [])]),
    "xl/sheets/one.xml": `<worksheet xmlns="${S}"><sheetData><row r="1">${str("B1", name, 0)}</row>${values.map((v, i) => `<row r="${i + 2}">${str(`A${i + 2}`, categories[i], i + 1)}<c r="B${i + 2}" ${error ? 't="e"' : ""}>${formula ? "<f>1/0</f>" : ""}<v>${escape(v)}</v></c></row>`).join("")}</sheetData></worksheet>`,
    ...(shared ? { "xl/strings.xml": `<sst xmlns="${S}">${[name, ...categories].map(v => `<si><t>${escape(v)}</t></si>`).join("")}</sst>` } : {}),
  };
  return pack(parts, contentTypes(undefined, Object.keys(parts).filter(p => !p.endsWith(".rels")).map(p => ({ PartName: `/${p}`, ContentType: `application/vnd.openxmlformats-officedocument.spreadsheetml.${p.endsWith("workbook.xml") ? "sheet.main" : p.endsWith("strings.xml") ? "sharedStrings" : "worksheet"}+xml` }))));
}
export function chartFixture({ xml = chartXml(), strict = false, content = frame(), book, external = false } = {}) {
  const f = contentFixture(content, strict);
  f.parts["deck/_rels/z.xml.rels"] = relationships([f.rel("layout", "slideLayout", "layout.xml"), f.rel("chart", "chart", "chart.xml")]);
  f.parts["deck/chart.xml"] = strict ? xml.replaceAll(CHART, "http://purl.oclc.org/ooxml/drawingml/chart").replaceAll("http://schemas.openxmlformats.org/drawingml/2006/main", f.A).replaceAll("http://schemas.openxmlformats.org/officeDocument/2006/relationships", f.R) : xml;
  if (strict) f.parts["deck/z.xml"] = f.parts["deck/z.xml"].replaceAll(CHART, "http://purl.oclc.org/ooxml/drawingml/chart");
  f.types["deck/chart.xml"] = "drawingml.chart";
  if (book || external) f.parts["deck/_rels/chart.xml.rels"] = relationships([{ ...f.rel("book", "package", external ? "https://example.invalid/no-network.xlsx" : "book.xlsx"), ...(external ? { TargetMode: "External" } : {}) }]);
  if (book) { f.parts["deck/book.xlsx"] = book; f.types["deck/book.xlsx"] = "spreadsheetml.sheet"; }
  return f;
}
// encodeStructure appends +xml; an embedded workbook is binary XLSX.
export function encodeChart(f) {
  return pack(f.parts, contentTypes(undefined, Object.entries(f.types).map(([part, type]) => ({ PartName: `/${part}`, ContentType: `application/vnd.openxmlformats-officedocument.${type}${part.endsWith(".xlsx") ? "" : "+xml"}` }))));
}
export const chartBytes = options => encodeChart(chartFixture(options));
