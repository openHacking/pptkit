import { placeholderFixture, inheritedShape, placeholderTransform } from "./placeholder-fixture.mjs";
import { relationships } from "./opc-fixture.mjs";

export const textBody = ({
  frame = "", fit = "", list = "",
  paragraphs = '<a:p><a:r><a:rPr lang="zh-CN"/><a:t>Inherited α😀</a:t></a:r></a:p>',
} = {}) => `<p:txBody><a:bodyPr ${frame}>${fit}</a:bodyPr><a:lstStyle>${list}</a:lstStyle>${paragraphs}</p:txBody>`;
export const masterList = '<a:defPPr><a:buChar char="*"/><a:defRPr sz="2400" b="1" u="sng" strike="sngStrike"><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:latin typeface="Parent Font"/></a:defRPr></a:defPPr><a:lvl1pPr algn="ctr" marL="25400" indent="-12700"><a:spcBef><a:spcPts val="150"/></a:spcBef></a:lvl1pPr>';
export const layoutList = '<a:lvl1pPr><a:defRPr sz="2000" i="0"><a:solidFill><a:srgbClr val="445566"/></a:solidFill></a:defRPr></a:lvl1pPr>';
export function setTextTheme(f, { accent1 = "123456", accent2 = "654321", major = "Major Theme Font", minor = "Minor Theme Font" } = {}) {
  const schemeColor = (value, name) => `<a:${name}><a:srgbClr val="${value}"/></a:${name}>`;
  const colors = [
    '<a:dk1><a:sysClr val="windowText" lastClr="101010"/></a:dk1>',
    schemeColor("FFFFFF", "lt1"), schemeColor("202020", "dk2"), schemeColor("EEEEEE", "lt2"),
    schemeColor(accent1, "accent1"), schemeColor(accent2, "accent2"),
    ...["334455", "445566", "556677", "667788"].map((value, i) => schemeColor(value, `accent${i + 3}`)),
    schemeColor("0000FF", "hlink"), schemeColor("800080", "folHlink"),
  ].join("");
  f.parts["deck/theme.xml"] = `<a:theme xmlns:a="${f.A}"><a:themeElements><a:clrScheme name="Test">${colors}</a:clrScheme><a:fontScheme name="Test"><a:majorFont><a:latin typeface="${major}"/></a:majorFont><a:minorFont><a:latin typeface="${minor}"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>`;
  return f;
}
export function addTextThemeOverride(f, owner, name, content) {
  const path = `deck/${name}.xml`, relPath = `deck/_rels/${owner}.xml.rels`;
  f.parts[path] = `<a:themeOverride xmlns:a="${f.A}">${content}</a:themeOverride>`;
  f.types[path] = "themeOverride";
  const relation = relationships([f.rel(`theme-${name}`, "themeOverride", `${name}.xml`)]);
  f.parts[relPath] = f.parts[relPath].replace("</Relationships>", relation.slice(relation.indexOf("<Relationship "), relation.indexOf("</Relationships>")) + "</Relationships>");
  return path;
}

export function setLayoutThemeFont(f, family, name = "font-override") {
  const scheme = /<a:fontScheme\b[\s\S]*?<\/a:fontScheme>/.exec(f.parts["deck/theme.xml"])[0];
  const override = scheme.replace(/(<a:majorFont><a:latin typeface=")[^"]*/, (_match, prefix) => prefix + family);
  addTextThemeOverride(f, "layout", name, override);
  return f;
}

export function textStyleFixture(strict = false) {
  const f = placeholderFixture(strict);
  const setMaster = (text, style = '<a:lvl1pPr><a:defRPr sz="3000" i="1"/></a:lvl1pPr>') =>
    f.set("deck/master.xml", "sldMaster", inheritedShape({ xfrm: placeholderTransform(), preset: "rect", text }),
      `<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="layout"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle>${style}</p:titleStyle><p:bodyStyle><a:lvl1pPr><a:defRPr sz="1800"/></a:lvl1pPr></p:bodyStyle><p:otherStyle/></p:txStyles>`);
  setMaster(textBody({ frame: 'lIns="12700" rIns="25400" tIns="38100" bIns="50800" anchor="ctr" wrap="square"', fit: '<a:normAutofit fontScale="75000"/>', list: masterList, paragraphs: '<a:p><a:r><a:rPr sz="9900"/><a:t>MASTER PROMPT MUST NOT COPY</a:t></a:r></a:p>' }));
  f.set("deck/layout.xml", "sldLayout", inheritedShape({ text: textBody({ frame: 'lIns="63500" tIns="0" anchor="b"', list: layoutList }) }));
  for (const [part, paragraph] of [["deck/z.xml", "Inherited α😀"], ["deck/a.xml", "Second page"]]) {
    f.set(part, "sld", inheritedShape({ text: textBody({ paragraphs: `<a:p><a:r><a:rPr lang="zh-CN"/><a:t>${paragraph}</a:t></a:r></a:p>` }) }));
  }
  return { ...f, setMaster };
}
