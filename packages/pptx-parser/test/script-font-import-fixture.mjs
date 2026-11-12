import { textStyleFixture, setTextTheme, textBody, masterList, layoutList } from "./text-style-fixture.mjs";
import { inheritedShape } from "./placeholder-fixture.mjs";

export const sourceFonts = (latin = "Latin Source", east = "Eastern Source", complex = "Complex Source") => `<a:latin typeface="${latin}"/><a:ea typeface="${east}"/><a:cs typeface="${complex}"/>`;
export function scriptFontImportFixture({ strict = false, run = sourceFonts(), list = "", paragraph = "", inherited = false } = {}) {
  const f = setTextTheme(textStyleFixture(strict), { major: "Heading Source", minor: "Body Source" });
  f.parts["deck/theme.xml"] = f.parts["deck/theme.xml"].replaceAll('<a:latin typeface="Heading Source"/>', sourceFonts("Heading Source", "Heading East", "Heading Complex"))
    .replaceAll('<a:latin typeface="Body Source"/>', sourceFonts("Body Source", "Body East", "Body Complex"));
  if (inherited) {
    f.parts["deck/master.xml"] = f.parts["deck/master.xml"].replace(masterList, `<a:defPPr><a:defRPr>${sourceFonts("Master Latin", "Master East", "Master Complex")}</a:defRPr></a:defPPr>`);
    f.parts["deck/layout.xml"] = f.parts["deck/layout.xml"].replace(layoutList, '<a:lvl1pPr><a:defRPr><a:latin typeface="Layout Latin"/></a:defRPr></a:lvl1pPr>');
  }
  for (const part of ["deck/z.xml", "deck/a.xml"]) f.set(part, "sld", inheritedShape({ text: textBody({ list, paragraphs: `<a:p>${paragraph ? `<a:pPr>${paragraph}</a:pPr>` : ""}<a:r><a:rPr>${run}</a:rPr><a:t>Latin中文العربية</a:t></a:r></a:p>` }) }));
  return f;
}

export function secondFontMaster(f) {
  f.parts["deck/main.data"] = f.parts["deck/main.data"].replace("</p:sldMasterIdLst>", '<p:sldMasterId id="2147483650" r:id="m2"/></p:sldMasterIdLst>');
  f.parts["deck/_rels/main.data.rels"] = f.parts["deck/_rels/main.data.rels"].replace("</Relationships>", `<Relationship Id="m2" Type="${f.R}/slideMaster" Target="master2.xml"/></Relationships>`);
  for (const kind of ["master", "layout", "theme"]) { f.parts[`deck/${kind}2.xml`] = f.parts[`deck/${kind}.xml`]; f.types[`deck/${kind}2.xml`] = f.types[`deck/${kind}.xml`]; }
  f.parts["deck/_rels/layout2.xml.rels"] = f.parts["deck/_rels/layout.xml.rels"].replace("master.xml", "master2.xml");
  f.parts["deck/_rels/master2.xml.rels"] = f.parts["deck/_rels/master.xml.rels"].replace("layout.xml", "layout2.xml").replace("theme.xml", "theme2.xml");
  f.parts["deck/_rels/a.xml.rels"] = f.parts["deck/_rels/a.xml.rels"].replace("layout.xml", "layout2.xml");
  return f;
}
