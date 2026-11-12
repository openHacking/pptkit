import test from "node:test";
import assert from "node:assert/strict";
import { importPptx } from "../dist/index.js";
import { encodeStructure } from "./structure-fixture.mjs";
import { relationships } from "./opc-fixture.mjs";
import { inheritedShape } from "./placeholder-fixture.mjs";
import { textBody, textStyleFixture, setTextTheme, layoutList, addTextThemeOverride as addOverride } from "./text-style-fixture.mjs";

const themedProperties = '<a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:latin typeface="+mj-lt"/>';
const localText = textBody({ paragraphs: `<a:p><a:r><a:rPr>${themedProperties}</a:rPr><a:t>Scoped theme</a:t></a:r></a:p>` });
const run = (session, page = 0) => session.document.slides[page].elements[0].text.content[0].runs[0];
const properties = (session, property, part = "/deck/z.xml") => session.report.textStyleInheritance.properties
  .filter(info => info.property === property && info.target.part === part);
const open = f => importPptx(encodeStructure(f), { projection: "content" });
const scheme = (f, name) => new RegExp(`<a:${name}\\b[\\s\\S]*?</a:${name}>`).exec(f.parts["deck/theme.xml"])[0];
const colorScheme = (f, value) => scheme(f, "clrScheme").replace("123456", value);
const fontScheme = (f, value) => scheme(f, "fontScheme").replace("Base Major", value);

function fixture(strict = false) {
  const f = setTextTheme(textStyleFixture(strict), { major: "Base Major", minor: "Base Minor" });
  for (const part of ["deck/z.xml", "deck/a.xml"]) f.set(part, "sld", inheritedShape({ text: localText }));
  return f;
}

test("partial color and font overrides inherit the absent component at slide, layout and master scope in both dialects", async () => {
  for (const strict of [false, true]) for (const owner of ["z", "layout", "master"]) for (const component of ["color", "font", "empty"]) {
    const f = fixture(strict);
    const path = addOverride(f, owner, "partial", component === "color" ? colorScheme(f, "ABCDEF")
      : component === "font" ? fontScheme(f, "Override Major") : "");
    const s = await open(f);
    try {
      assert.equal(run(s).style.color, component === "color" ? "ABCDEF" : "123456", `${strict}/${owner}/${component}`);
      assert.equal(run(s).style.fontFamily, component === "font" ? "Override Major" : "Base Major");
      assert.equal(properties(s, "run.color")[0].resolvedFrom.at(-1).part, component === "color" ? `/${path}` : "/deck/theme.xml");
      assert.equal(properties(s, "run.fontFamily")[0].resolvedFrom.at(-1).part, component === "font" ? `/${path}` : "/deck/theme.xml");
      if (owner === "z") {
        assert.equal(run(s, 1).style.color, "123456");
        assert.equal(run(s, 1).style.fontFamily, "Base Major");
      } else {
        assert.equal(run(s, 1).style.color, run(s).style.color);
        assert.equal(run(s, 1).style.fontFamily, run(s).style.fontFamily);
      }
    } finally {
      s.dispose();
    }
  }
});

test("each component selects slide then layout then master without depending on the declaration's owner", async () => {
  for (const inherited of [false, true]) {
    const f = fixture();
    if (inherited) {
      f.parts["deck/layout.xml"] = f.parts["deck/layout.xml"].replace(layoutList,
        `<a:lvl1pPr><a:defRPr>${themedProperties}</a:defRPr></a:lvl1pPr>`);
      for (const part of ["deck/z.xml", "deck/a.xml"]) f.set(part, "sld", inheritedShape({ text: textBody() }));
    }
    addOverride(f, "master", "master-override", colorScheme(f, "AA0000") + fontScheme(f, "Master Major"));
    addOverride(f, "layout", "layout-override", colorScheme(f, "00BB00"));
    addOverride(f, "z", "slide-override", fontScheme(f, "Slide Major"));
    const s = await open(f);
    try {
      assert.equal(run(s).style.color, "00BB00");
      assert.equal(run(s).style.fontFamily, "Slide Major");
      assert.equal(run(s, 1).style.color, "00BB00");
      assert.equal(run(s, 1).style.fontFamily, "Master Major");
      assert.equal(properties(s, "run.color")[0].source.part, inherited ? "/deck/layout.xml" : "/deck/z.xml");
      assert.equal(properties(s, "run.color")[0].resolvedFrom.at(-1).part, "/deck/layout-override.xml");
      assert.equal(properties(s, "run.fontFamily")[0].resolvedFrom.at(-1).part, "/deck/slide-override.xml");
      assert.equal(properties(s, "run.fontFamily", "/deck/a.xml")[0].resolvedFrom.at(-1).part, "/deck/master-override.xml");
    } finally {
      s.dispose();
    }
  }
});

test("present schemes with unsupported or missing slots block older schemes instead of merging individual slots", async () => {
  for (const content of [
    '<a:clrScheme name="Missing slot"><a:dk1><a:srgbClr val="101010"/></a:dk1></a:clrScheme>',
    '<a:clrScheme name="Unsupported"><a:accent1><a:srgbClr val="FFFFFF"><a:hueOff val="25000"/></a:srgbClr></a:accent1></a:clrScheme>',
    '<a:fontScheme name="Empty font"><a:majorFont><a:latin typeface=""/></a:majorFont></a:fontScheme>',
    '<a:fontScheme name="Missing font"/>',
  ]) {
    const f = fixture(), color = content.includes("clrScheme");
    addOverride(f, "z", "unsupported", content);
    const s = await open(f);
    try {
      const property = color ? "run.color" : "run.fontFamily";
      assert.equal(properties(s, property)[0].status, "unsupported");
      assert.notDeepEqual(run(s).style[color ? "color" : "fontFamily"], color ? "123456" : "Base Major");
      assert.equal(run(s).style[color ? "fontFamily" : "color"], color ? "Base Major" : "123456");
      assert.equal(run(s, 1).style.color, "123456");
      assert.equal(run(s, 1).style.fontFamily, "Base Major");
    } finally {
      s.dispose();
    }
  }
});

test("an unsupported higher component blocks a valid layout component while independent components still inherit", async () => {
  const f = fixture();
  addOverride(f, "layout", "layout-override", colorScheme(f, "00BB00") + fontScheme(f, "Layout Major"));
  addOverride(f, "z", "slide-override", '<a:clrScheme name="Unsupported"><a:accent1><a:schemeClr val="accent2"/></a:accent1></a:clrScheme>');
  const s = await open(f);
  try {
    assert.equal(properties(s, "run.color")[0].status, "unsupported");
    assert.notEqual(run(s).style.color, "00BB00");
    assert.equal(run(s).style.fontFamily, "Layout Major");
    assert.equal(run(s, 1).style.color, "00BB00");
  } finally {
    s.dispose();
  }
});

test("a shared partial override never mixes the inherited components of two different masters", async () => {
  for (const component of ["color", "font"]) {
    const f = fixture();
    f.parts["deck/main.data"] = f.parts["deck/main.data"].replace("</p:sldMasterIdLst>", '<p:sldMasterId id="2147483650" r:id="m2"/></p:sldMasterIdLst>');
    f.parts["deck/_rels/main.data.rels"] = relationships([f.rel("a", "slide", "a.xml"), f.rel("m", "slideMaster", "master.xml"), f.rel("z", "slide", "z.xml"), f.rel("m2", "slideMaster", "master2.xml")]);
    for (const kind of ["layout", "master", "theme"]) {
      f.parts[`deck/${kind}2.xml`] = f.parts[`deck/${kind}.xml`];
      f.types[`deck/${kind}2.xml`] = f.types[`deck/${kind}.xml`];
    }
    f.parts["deck/theme2.xml"] = f.parts["deck/theme2.xml"].replace("123456", "789ABC").replace("Base Major", "Second Major");
    f.parts["deck/_rels/layout2.xml.rels"] = relationships([f.rel("master", "slideMaster", "master2.xml")]);
    f.parts["deck/_rels/master2.xml.rels"] = relationships([f.rel("layout", "slideLayout", "layout2.xml"), f.rel("theme", "theme", "theme2.xml")]);
    f.parts["deck/_rels/a.xml.rels"] = relationships([f.rel("layout", "slideLayout", "layout2.xml")]);
    addOverride(f, "layout", "shared", component === "color" ? colorScheme(f, "A0B0C0") : fontScheme(f, "Shared Major"));
    f.parts["deck/_rels/layout2.xml.rels"] = relationships([f.rel("master", "slideMaster", "master2.xml"), f.rel("override", "themeOverride", "shared.xml")]);
    const s = await open(f);
    try {
      assert.deepEqual([run(s).style.color, run(s, 1).style.color], component === "color" ? ["A0B0C0", "A0B0C0"] : ["123456", "789ABC"]);
      assert.deepEqual([run(s).style.fontFamily, run(s, 1).style.fontFamily], component === "font" ? ["Shared Major", "Shared Major"] : ["Base Major", "Second Major"]);
    } finally { s.dispose(); }
  }
});

test("theme overrides use the effective MCE component tree and record the chosen fallback source", async () => {
  const f = fixture();
  const path = addOverride(f, "layout", "mce-override", `<mc:AlternateContent><mc:Choice Requires="future"><future:font/></mc:Choice><mc:Fallback>${fontScheme(f, "MCE Major")}</mc:Fallback></mc:AlternateContent>`);
  f.parts[path] = f.parts[path].replace("<a:themeOverride ", '<a:themeOverride xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:future="urn:pptkit:test:future" mc:Ignorable="future" ');
  const s = await open(f);
  try {
    assert.equal(run(s).style.color, "123456");
    assert.equal(run(s).style.fontFamily, "MCE Major");
    assert.equal(properties(s, "run.fontFamily")[0].resolvedFrom[0].part, `/${path}`);
    assert.ok(s.report.inspection.markupCompatibility.some(info => info.part === `/${path}`));
  } finally {
    s.dispose();
  }
});

test("override relationships require the matching content type and root in the source dialect", async () => {
  for (const corruption of ["type", "root", "namespace", "duplicate", "external", "duplicate-scheme"]) {
    const f = fixture(), path = addOverride(f, "z", "invalid", fontScheme(f, "Override Major"));
    if (corruption === "type") f.types[path] = "theme";
    if (corruption === "root") f.parts[path] = f.parts[path].replaceAll("a:themeOverride", "a:theme");
    if (corruption === "namespace") f.parts[path] = f.parts[path].replace(f.A, "urn:pptkit:test:wrong");
    if (corruption === "duplicate") f.parts["deck/_rels/z.xml.rels"] = f.parts["deck/_rels/z.xml.rels"]
      .replace("</Relationships>", `<Relationship Id="second-override" Type="${f.R}/themeOverride" Target="invalid.xml"/></Relationships>`);
    if (corruption === "external") f.parts["deck/_rels/z.xml.rels"] = f.parts["deck/_rels/z.xml.rels"]
      .replace('Target="invalid.xml"', 'Target="https://example.invalid/theme.xml" TargetMode="External"');
    if (corruption === "duplicate-scheme") f.parts[path] = f.parts[path]
      .replace("</a:themeOverride>", `${fontScheme(f, "Duplicate Major")}</a:themeOverride>`);
    await assert.rejects(open(f), error => error.code === "import-invalid-package"
      && error.cause?.code === "presentation-invalid", corruption);
  }
});

test("component selection remains isolated across imports and original-byte no-op saves retain partial overrides", async () => {
  const f = fixture(), path = addOverride(f, "layout", "shared", fontScheme(f, "First Major"));
  const input = encodeStructure(f), first = await importPptx(input, { projection: "content" });
  f.parts[path] = f.parts[path].replace("First Major", "Second Major");
  const second = await open(f);
  try {
    assert.equal(run(first).style.fontFamily, "First Major");
    assert.equal(run(second).style.fontFamily, "Second Major");
    const plan = await first.prepareSave({ mode: "no-op" });
    assert.equal(plan.status, "ready");
    assert.deepEqual((await first.save(plan.plan)).bytes, new Uint8Array(input));
    assert.ok(Object.isFrozen(properties(first, "run.fontFamily")[0].resolvedFrom[0]));
    assert.equal((await first.prepareSave({ mode: "edit" })).status, "blocked");
  } finally {
    first.dispose();
    second.dispose();
  }
});
