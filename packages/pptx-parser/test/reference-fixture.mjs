import { contentFixture, transform } from "./content-fixture.mjs";
import { relationships } from "./opc-fixture.mjs";

export const click = (id = "url", attributes = "", content = "") => `<a:hlinkClick r:id="${id}" ${attributes}>${content}</a:hlinkClick>`;
export const withClick = (drawing, link = click()) => drawing.replace(/(<p:cNvPr\b[^>]*?)\/>/, `$1>${link}</p:cNvPr>`);
export const connector = ({ id = 2, preset = "straightConnector1", startId, startSite = 3, endId, endSite = 1, xfrm = transform(), props = "", nv = "" } = {}) =>
  `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="${id}" name="Connector ${id}"/><p:cNvCxnSpPr>${startId === undefined ? "" : `<a:stCxn id="${startId}" idx="${startSite}"/>`}${endId === undefined ? "" : `<a:endCxn id="${endId}" idx="${endSite}"/>`}${nv}</p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr><p:spPr>${xfrm}<a:prstGeom prst="${preset}"><a:avLst/></a:prstGeom><a:ln w="25400"><a:solidFill><a:srgbClr val="123456"/></a:solidFill><a:prstDash val="dashDot"/><a:headEnd type="oval"/><a:tailEnd type="triangle"/></a:ln>${props}</p:spPr></p:cxnSp>`;
export const group = (content, id = 10, origin = 5) => `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${id}" name="Group ${id}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm rot="5400000" flipH="1"><a:off x="1270000" y="1270000"/><a:ext cx="2540000" cy="1270000"/><a:chOff x="${origin * 12700}" y="${origin * 12700}"/><a:chExt cx="1270000" cy="635000"/></a:xfrm></p:grpSpPr>${content}</p:grpSp>`;
export function referenceFixture(content, strict = false, extras = []) {
  const f = contentFixture(content, strict);
  f.parts["deck/_rels/z.xml.rels"] = relationships([
    f.rel("layout", "slideLayout", "layout.xml"),
    { ...f.rel("url", "hyperlink", "https://example.invalid/a?x=1&y=%E4%B8%AD#part"), TargetMode: "External" },
    f.rel("jump", "slide", "a.xml"), f.rel("self", "slide", "z.xml"),
    ...extras.map(value => typeof value === "function" ? value(f) : value),
  ]);
  return f;
}
