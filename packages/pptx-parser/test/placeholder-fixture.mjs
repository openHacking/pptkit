import { structureFixture, encodeStructure } from "./structure-fixture.mjs";
import { shape, body } from "./content-fixture.mjs";

export const inheritedShape = ({ id = "2", ph = 'type="title"', xfrm = "", preset = "", text = "", fill = "<a:noFill/>", props = "", ...options } = {}) => {
  let result = shape({ id, xfrm, text, fill, props, ...options }).replace("<p:nvPr/>", `<p:nvPr><p:ph ${ph}/></p:nvPr>`);
  if (!preset) result = result.replace('<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>', "");
  else result = result.replace('prst="rect"', `prst="${preset}"`);
  return result;
};
export const placeholderTransform = ({ x = 127000, y = 254000, width = 1270000, height = 635000, attrs = "", off = true, ext = true } = {}) =>
  `<a:xfrm ${attrs}>${off ? `<a:off x="${x}" y="${y}"/>` : ""}${ext ? `<a:ext cx="${width}" cy="${height}"/>` : ""}</a:xfrm>`;
export function placeholderFixture(strict = false) {
  const f = structureFixture(strict);
  const tree = content => `<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name="Root"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${content}</p:spTree>`;
  const set = (part, name, content, tail = "", attrs = "") => {
    f.parts[part] = f.p(name, `<p:cSld>${tree(content)}</p:cSld>${tail}`, `xmlns:a="${f.A}" ${attrs}`);
  };
  set("deck/z.xml", "sld", inheritedShape({ text: body() }), "", 'show="0"');
  set("deck/a.xml", "sld", inheritedShape({ text: body() }));
  set("deck/layout.xml", "sldLayout", inheritedShape({ xfrm: placeholderTransform({ ext: false }) }));
  set("deck/master.xml", "sldMaster", inheritedShape({ xfrm: placeholderTransform({ width: 2540000, height: 1270000 }), preset: "rect", text: body('<a:p><a:r><a:t>MASTER PROMPT MUST NOT COPY</a:t></a:r></a:p>') }),
    '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="layout"/></p:sldLayoutIdLst>');
  return { ...f, tree, set };
}
export const placeholderBytes = (strict = false) => encodeStructure(placeholderFixture(strict));
