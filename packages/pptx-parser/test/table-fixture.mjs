// Independent DrawingML fixture: does not call PPTKit's table writer.
import { contentFixture } from "./content-fixture.mjs";
import { encodeStructure } from "./structure-fixture.mjs";

export const tableBody = (paragraphs = "<a:p/>", attributes = "") =>
  `<a:txBody><a:bodyPr ${attributes}/><a:lstStyle/>${paragraphs}</a:txBody>`;
export const cell = ({ attributes = "", text = "", body = tableBody(`<a:p><a:r><a:t>${text}</a:t></a:r></a:p>`), properties = "<a:tcPr/>" } = {}) =>
  `<a:tc ${attributes}>${body}${properties}</a:tc>`;
export const tableRow = (cells, height = 20) => `<a:tr h="${height * 12700}">${cells.join("")}</a:tr>`;
export const tableTransform = (attributes = "", width = 100, height = 50) =>
  `<p:xfrm ${attributes}><a:off x="127000" y="254000"/><a:ext cx="${width * 12700}" cy="${height * 12700}"/></p:xfrm>`;
export function table({
  id = "2", columns = [40, 60],
  rows = [tableRow([cell({ text: "R1C1" }), cell({ text: "R1C2" })]), tableRow([cell({ text: "R2C1" }), cell({ text: "R2C2" })], 30)],
  properties = "<a:tblPr/>", transform = tableTransform(), extra = "",
  uri = "http://schemas.openxmlformats.org/drawingml/2006/table",
} = {}) {
  return `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="Table ${id}" descr="Table description"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>${transform}<a:graphic><a:graphicData uri="${uri}"><a:tbl>${properties}<a:tblGrid>${columns.map(w => `<a:gridCol w="${w * 12700}"/>`).join("")}</a:tblGrid>${rows.join("")}${extra}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;
}
export const tableFixture = (markup = table(), strict = false) => contentFixture(markup, strict);
export const tableBytes = (markup, strict) => encodeStructure(tableFixture(markup, strict));
