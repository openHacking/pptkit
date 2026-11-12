import { structureFixture, encodeStructure } from "./structure-fixture.mjs";
export const transform = (extra = "") => `<a:xfrm ${extra}><a:off x="127000" y="254000"/><a:ext cx="1270000" cy="635000"/></a:xfrm>`;
export const body = (paragraphs = '<a:p><a:r><a:rPr sz="2400" b="1" lang="zh-CN"><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:latin typeface="Test Font"/></a:rPr><a:t> A &amp; B </a:t></a:r></a:p>') => `<p:txBody><a:bodyPr lIns="12700" rIns="25400" tIns="38100" bIns="50800" anchor="ctr" wrap="none"><a:noAutofit/></a:bodyPr><a:lstStyle/>${paragraphs}</p:txBody>`;
export const shape = ({ id = "2", preset = "rect", props = "", text = "", nv = "", txBox = false, xfrm = transform(), fill = "<a:noFill/>" } = {}) => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Object ${id}" descr="Description" ${nv}/><p:cNvSpPr txBox="${txBox ? 1 : 0}"/><p:nvPr/></p:nvSpPr><p:spPr>${xfrm}<a:prstGeom prst="${preset}"><a:avLst/></a:prstGeom>${fill}${props}</p:spPr>${text}</p:sp>`;
export function contentFixture(content = shape({ txBox: true, text: body() }), strict = false) {
  const f = structureFixture(strict);
  f.parts["deck/z.xml"] = f.p("sld", `<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name="Root"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${content}</p:spTree></p:cSld>`, `xmlns:a="${f.A}" show="0"`);
  return f;
}
export const bytes = (content, strict) => encodeStructure(contentFixture(content, strict));
