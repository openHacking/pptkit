import test from "node:test";
import assert from "node:assert/strict";
import { importPptx, describeImportCapabilities } from "../dist/index.js";
import { importPptxFile } from "../dist/node.js";
import { mkdtemp, writeFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { structureFixture } from "./structure-fixture.mjs";
import { pack, contentTypes, relationships } from "./opc-fixture.mjs";
import { png } from "./image-fixture.mjs";
import { normalizePresentation } from "@pptkit/core";

test("artwork and background capability references retain actual reported tiers", async () => {
  for (const strict of [false, true]) {
    const session = await importPptx(fixture(strict, false, "backgrounds"), { projection: "content", masterArtwork: "artwork-v6" });
    try {
      for (const feature of ["master-artwork.shape", "artwork.background"]) {
        const findings = session.report.findings.filter(f => f.feature === feature);
        assert.ok(findings.length, feature);
        const view = describeImportCapabilities(session.report, { filter: { feature } });
        assert.ok(view.subjects.length);
        for (const subject of view.subjects) {
          assert.ok(subject.references.tasks.includes("IMP-05.1.3.3.3"));
          const row = subject.operations.find(operation => operation.operation === "import");
          assert.ok(row.observations.every(observation => findings.some(f => f.id === observation.id && f.tier === observation.tier)));
          assert.equal(subject.operations.find(operation => operation.operation === "edit-save").status, "unreported");
        }
        assert.equal(view.candidateSaveValidation, "not-performed");
      }
    } finally { session.dispose(); }
  }
});

test('shared master and layout picture imports retain decorative flags', async () => {
  for (const strict of [false, true]) {
    const bytes = fixture(strict, false, 'decorative'), session = await importPptx(bytes, { projection: 'content', masterArtwork: 'pictures-v1' });
    try {
      assert.equal(session.document.masters[0].elements[0].accessibility.decorative, true);
      assert.equal(session.document.layouts[0].elements[0].accessibility.decorative, false);
      const plan = await session.prepareSave({ mode: 'no-op' }); assert.equal(plan.status, 'ready');
      assert.deepEqual((await session.save(plan.plan)).bytes, new Uint8Array(bytes));
    } finally { session.dispose(); }
  }
});

test("file imports capture every master artwork profile before I/O and match byte imports", async t => {
  const dir = await mkdtemp(path.join(tmpdir(), "pptkit-master-options-"));
  t.after(() => rm(dir, { recursive: true, force: true }));
  const profiles = ["pictures-v1", ...Array.from({ length: 11 }, (_, index) => `artwork-v${index + 1}`)];
  for (const strict of [false, true]) {
    const bytes = fixture(strict), file = path.join(dir, `${strict}.pptx`); await writeFile(file, bytes);
    for (const masterArtwork of profiles) {
      const options = { projection: "content", masterArtwork }, expected = await importPptx(bytes, options);
      const pending = importPptxFile(file, options); options.masterArtwork = "invalid";
      const actual = await pending;
      try {
        assert.deepEqual(normalizePresentation(actual.session.document), normalizePresentation(expected.document), masterArtwork);
        assert.deepEqual(actual.session.report, expected.report, masterArtwork);
        assert.ok(actual.session.document.masters.some(master => master.elements.length));
      } finally { actual.session.dispose(); expected.dispose(); }
    }
  }
});

function fixture(strict = false, unlisted = false, variant = "") {
  const f = structureFixture(strict);
  const picture = id => `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="Picture ${id}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="image"/><a:srcRect l="10000"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="12700" y="25400"/><a:ext cx="127000" cy="254000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
  const group = (id, body, x, y, childX, childY, childWidth) => `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${id}" name="Group ${id}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm rot="60000" flipH="1"><a:off x="${x*12700}" y="${y*12700}"/><a:ext cx="508000" cy="381000"/><a:chOff x="${childX*12700}" y="${childY*12700}"/><a:chExt cx="${childWidth*12700}" cy="254000"/></a:xfrm></p:grpSpPr>${body}</p:grpSp>`;
  const shape = id => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Shape ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="254000" cy="254000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill><a:ln w="12700"><a:noFill/></a:ln></p:spPr></p:sp>`;
  const tree = child => `<p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>${child}</p:spTree></p:cSld>`;
  f.parts["deck/master.xml"] = `<p:sldMaster xmlns:p="${f.P}" xmlns:a="${f.A}" xmlns:r="${f.R}">${tree(picture(3) + '<p:sp/>')}<p:sldLayoutIdLst>${unlisted ? "" : '<p:sldLayoutId id="2147483649" r:id="layout"/>'}</p:sldLayoutIdLst></p:sldMaster>`;
  if (variant === "placeholder") f.parts["deck/master.xml"] = f.parts["deck/master.xml"].replace(picture(3),picture(3).replace('<p:nvPr/>','<p:nvPr><p:ph type="pic"/></p:nvPr>'));
  if (variant === "group") f.parts["deck/master.xml"] = f.parts["deck/master.xml"].replace(picture(3),`<p:grpSp>${picture(3)}</p:grpSp>`);
  if (variant === "nested" || variant === "zero-group") f.parts["deck/master.xml"] = f.parts["deck/master.xml"].replace(picture(3),group(10,group(11,picture(3),4,3,1,1,20),10,20,2,1,variant === "nested" ? 100 : 0));
  if (variant === "duplicate") f.parts["deck/master.xml"] = f.parts["deck/master.xml"].replace(picture(3),group(3,picture(3),10,20,2,1,0));
  if (variant === "bad-id") f.parts["deck/master.xml"] = f.parts["deck/master.xml"].replace('id="3" name="Picture 3"','id="invalid" name="Picture 3"');
  if (variant === "shapes") f.parts["deck/master.xml"] = f.parts["deck/master.xml"].replace(picture(3),shape(8)+group(10,shape(9),10,20,2,1,100)+picture(3));
  if (variant === "shape-text") f.parts["deck/master.xml"] = f.parts["deck/master.xml"].replace(picture(3),shape(8).replace('</p:sp>','<p:txBody/></p:sp>')+picture(3));
  if (variant === "rich-text") { const body='<p:txBody><a:bodyPr lIns="12700"/><a:lstStyle/><a:p><a:r><a:rPr sz="2400" b="1"><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:hlinkClick r:id="unmapped"/></a:rPr><a:t>你好 &amp; test</a:t></a:r><a:br/><a:r><a:t>second</a:t></a:r><a:fld id="{00000000-0000-0000-0000-000000000001}" type="slidenum"><a:t>7</a:t></a:fld></a:p><a:p><a:endParaRPr/></a:p></p:txBody>'; const labelled=id=>shape(id).replace('</p:sp>',body+'</p:sp>');f.parts["deck/master.xml"]=f.parts["deck/master.xml"].replace(picture(3),labelled(8)+group(10,labelled(9),10,20,2,1,100)+picture(3)); }
  if (variant === "shape-theme") f.parts["deck/master.xml"] = f.parts["deck/master.xml"].replace(picture(3),shape(8).replace('<a:srgbClr val="0000FF"/>','<a:schemeClr val="accent1"/>')+picture(3));
  if (variant === "shape-adjust") f.parts["deck/master.xml"] = f.parts["deck/master.xml"].replace(picture(3),shape(8).replace('<a:avLst/>','<a:avLst><a:gd name="adj" fmla="val 1"/></a:avLst>')+picture(3));
  if (variant === "shape-no-line") f.parts["deck/master.xml"] = f.parts["deck/master.xml"].replace(picture(3),shape(8).replace('<a:ln w="12700"><a:noFill/></a:ln>','')+picture(3));
  if (variant === "rgb-transform") f.parts["deck/master.xml"] = f.parts["deck/master.xml"].replace(picture(3),shape(8).replace('<a:srgbClr val="0000FF"/>','<a:srgbClr val="0000FF"><a:lumMod val="50000"/><a:alpha val="50000"/></a:srgbClr>')+picture(3));
  if (variant === "depth48" || variant === "depth49") { let nested=picture(3); for(let i=0;i<(variant === "depth48" ? 48 : 49);i++) nested=group(100+i,nested,1,1,0,0,20); f.parts["deck/master.xml"]=f.parts["deck/master.xml"].replace(picture(3),nested); }
  f.parts["deck/layout.xml"] = `<p:sldLayout xmlns:p="${f.P}" xmlns:a="${f.A}" xmlns:r="${f.R}" showMasterSp="false">${tree(picture(4))}</p:sldLayout>`;
  if (variant === "shape-gradient" || variant === "shape-gradient-unsupported") {
    const fill=`<a:gradFill${variant === "shape-gradient-unsupported" ? ' flip="x"' : ''}><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:lin ang="2700000" scaled="1"/></a:gradFill>`;
    const gradientShape=id=>shape(id).replace('<a:solidFill><a:srgbClr val="0000FF"/></a:solidFill>',fill);
    f.parts["deck/master.xml"]=f.parts["deck/master.xml"].replace(picture(3),gradientShape(8)+group(10,gradientShape(9),10,20,2,1,100)+picture(3));
    f.parts["deck/layout.xml"]=f.parts["deck/layout.xml"].replace(picture(4),gradientShape(12)+picture(4));
  }
  if (variant === "rich-text") f.parts["deck/layout.xml"]=f.parts["deck/layout.xml"].replace(picture(4),shape(12).replace('</p:sp>','<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Layout text</a:t></a:r></a:p></p:txBody></p:sp>')+picture(4));
  if (["gradient-theme","gradient-theme-missing","theme-text","theme-missing","theme-map","theme-paint","paint-unsupported","paint-missing","line-unsupported","backgrounds","background-unsupported","background-matrix","background-matrix-missing","background-matrix-alpha","background-matrix-zero","background-matrix-gradient","background-matrix-unbound"].includes(variant)) {
    const body=`<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr><a:solidFill><a:schemeClr val="${variant === "theme-missing" ? "accent2" : "accent1"}"/></a:solidFill><a:latin typeface="+mj-lt"/></a:rPr><a:t>Themed</a:t></a:r></a:p></p:txBody>`;
    const labelled=id=>{let result=shape(id);if(["theme-paint","paint-unsupported","paint-missing","line-unsupported"].includes(variant)) result=result.replace('<a:srgbClr val="0000FF"/>',`<a:schemeClr val="${variant === "paint-missing" ? "accent2" : "accent1"}"><a:${variant === "paint-unsupported" ? "hueOff" : "alpha"} val="50000"/></a:schemeClr>`).replace('<a:ln w="12700"><a:noFill/></a:ln>',`<a:ln w="12700"><a:solidFill><a:schemeClr val="accent1">${variant === "line-unsupported" ? '<a:hueOff val="50000"/>' : ''}</a:schemeClr></a:solidFill></a:ln>`);return result.replace('</p:sp>',body+'</p:sp>');};
    f.parts["deck/master.xml"]=f.parts["deck/master.xml"].replace(picture(3),labelled(8)+picture(3));
    f.parts["deck/layout.xml"]=f.parts["deck/layout.xml"].replace(picture(4),labelled(12)+picture(4));
    f.parts["deck/theme.xml"]=`<a:theme xmlns:a="${f.A}"><a:themeElements><a:clrScheme name="base"><a:accent1><a:srgbClr val="123456"/></a:accent1></a:clrScheme><a:fontScheme name="base"><a:majorFont><a:latin typeface="Base Font"/></a:majorFont></a:fontScheme></a:themeElements></a:theme>`;
    f.parts["deck/override.xml"]=`<a:themeOverride xmlns:a="${f.A}"><a:clrScheme name="override"><a:accent1><a:srgbClr val="AA0000"/></a:accent1></a:clrScheme><a:fontScheme name="override"><a:majorFont><a:latin typeface="Layout Font"/></a:majorFont></a:fontScheme></a:themeOverride>`;
    f.parts["deck/slide-override.xml"]=`<a:themeOverride xmlns:a="${f.A}"><a:clrScheme name="slide"><a:accent1><a:srgbClr val="FFFFFF"/></a:accent1></a:clrScheme><a:fontScheme name="slide"><a:majorFont><a:latin typeface="Slide Font"/></a:majorFont></a:fontScheme></a:themeOverride>`;
    f.types["deck/slide-override.xml"]="themeOverride";f.parts["deck/_rels/z.xml.rels"]=relationships([f.rel("layout","slideLayout","layout.xml"),f.rel("override","themeOverride","slide-override.xml")]);
    if (["backgrounds","background-unsupported","background-matrix","background-matrix-missing","background-matrix-alpha","background-matrix-zero","background-matrix-gradient","background-matrix-unbound"].includes(variant)) {
      const matrix=variant.startsWith("background-matrix"), index=variant === "background-matrix-missing" ? 1002 : variant === "background-matrix-zero" ? 1000 : 1001;
      const bg=matrix ? `<p:bg><p:bgRef idx="${index}">${variant === "background-matrix-unbound" ? "" : `<a:schemeClr val="accent1">${variant === "background-matrix-alpha" ? '<a:alpha val="40000"/>' : ''}</a:schemeClr>`}</p:bgRef></p:bg>` : '<p:bg><p:bgPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></p:bgPr></p:bg>';
      if(matrix) {const selected=variant === "background-matrix-gradient" ? '<a:gradFill/>' : '<a:solidFill><a:schemeClr val="phClr"><a:alphaMod val="50000"/></a:schemeClr></a:solidFill>';f.parts["deck/theme.xml"]=f.parts["deck/theme.xml"].replace('</a:themeElements>',`<a:fmtScheme name="matrix"><a:bgFillStyleLst>${selected}</a:bgFillStyleLst></a:fmtScheme></a:themeElements>`);}
      f.parts["deck/master.xml"]=f.parts["deck/master.xml"].replace('<p:cSld>','<p:cSld>'+bg);
      f.parts["deck/layout.xml"]=f.parts["deck/layout.xml"].replace('<p:cSld>','<p:cSld>'+bg);
      const slideBg=variant === "background-unsupported" ? '<p:bg><p:bgRef idx="1001"><a:schemeClr val="accent1"/></p:bgRef></p:bg>' : bg;
      f.parts["deck/z.xml"]=f.parts["deck/z.xml"].replace('<p:sld ',`<p:sld xmlns:a="${f.A}" `).replace('<p:cSld name="First"/>','<p:cSld name="First">'+slideBg+'</p:cSld>');
    }
    if(variant === "theme-map") {f.parts["deck/override.xml"]=f.parts["deck/override.xml"].replace('</a:clrScheme>','<a:accent2><a:srgbClr val="00CC00"/></a:accent2></a:clrScheme>');f.parts["deck/layout.xml"]=f.parts["deck/layout.xml"].replace('</p:sldLayout>','<p:clrMapOvr><a:overrideClrMapping accent1="accent2"/></p:clrMapOvr></p:sldLayout>');}
    f.types["deck/override.xml"]="themeOverride";
  }
  if (variant.startsWith("slide-pattern-")) {
    const kind=variant.slice("slide-pattern-".length),preset=kind === "unknown" ? "smCheck" : kind.startsWith("pct") ? kind : kind === "vert" ? "vert" : kind === "cross" ? "cross" : kind === "dnDiag" ? "dnDiag" : kind === "upDiag" ? "upDiag" : kind === "diagCross" ? "diagCross" : "horz";
    const bg=kind === "missing" ? "" : '<a:bgClr><a:srgbClr val="0000FF"/></a:bgClr>';
    const fg=kind === "unresolved" ? '<a:schemeClr val="accent9"/>' : kind === "transform" ? '<a:srgbClr val="FF0000"><a:hueOff val="50000"/></a:srgbClr>' : `<a:srgbClr val="FF0000"><a:alpha val="${kind === "zero" ? "0" : "50000"}"/></a:srgbClr>`;
    const attr=kind === "no-preset" ? "" : ` prst="${preset}"`;
    const fill=`<a:pattFill${attr}><a:fgClr>${fg}</a:fgClr>${bg}</a:pattFill>`;
    const entry=shape(8).replace('<a:solidFill><a:srgbClr val="0000FF"/></a:solidFill>',fill);
    f.parts["deck/z.xml"]=f.parts["deck/z.xml"].replace('<p:sld ',`<p:sld xmlns:a="${f.A}" `).replace('<p:cSld name="First"/>',`<p:cSld name="First"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${entry}</p:spTree></p:cSld>`);
  }
  if (variant.startsWith("slide-shadow-")) {
    const kind=variant.slice("slide-shadow-".length),extra=kind === "scale" ? ' sx="90000"' : '',rotation=kind === "rotation" ? '0' : '1',direction=kind === "missing" ? '' : ' dir="5400000"',color=kind === "color" ? '<a:schemeClr val="accent9"/>' : '<a:srgbClr val="000000"><a:alpha val="50000"/></a:srgbClr>',other=kind === "multiple" ? '<a:glow/>' : '';
    const effect=`<a:effectLst><a:outerShdw${kind === "no-blur" ? "" : ' blurRad="76200"'} dist="101600"${direction} algn="ctr" rotWithShape="${rotation}"${extra}>${color}</a:outerShdw>${other}</a:effectLst>`;
    const entry=shape(8).replace('</p:spPr>',`${effect}</p:spPr>`);
    f.parts["deck/z.xml"]=f.parts["deck/z.xml"].replace('<p:sld ',`<p:sld xmlns:a="${f.A}" `).replace('<p:cSld name="First"/>',`<p:cSld name="First"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${entry}</p:spTree></p:cSld>`);
  }
  if (variant === "slide-gradient" || variant === "slide-gradient-flip") {
    const fill=`<a:gradFill${variant === "slide-gradient-flip" ? ' flip="x"' : ''}><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>`;
    const entry=shape(8).replace('<a:solidFill><a:srgbClr val="0000FF"/></a:solidFill>',fill);
    f.parts["deck/z.xml"]=f.parts["deck/z.xml"].replace('<p:sld ',`<p:sld xmlns:a="${f.A}" `).replace('<p:cSld name="First"/>',`<p:cSld name="First"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${entry}</p:spTree></p:cSld>`);
  }
  if (variant.startsWith("gradient-")) {
    const first = '<a:gs pos="0"><a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr></a:gs>';
    const second = `<a:gs pos="${variant === "gradient-unsorted" ? "0" : "100000"}"><a:srgbClr val="0000FF"/></a:gs>`;
    let stops = variant === "gradient-unsorted" ? second.replace('pos="0"','pos="100000"') + first : first + second;
    if (variant === "gradient-theme" || variant === "gradient-theme-missing") stops=`<a:gs pos="0"><a:schemeClr val="${variant === "gradient-theme" ? "accent1" : "accent2"}"><a:alpha val="80000"/><a:alphaMod val="50000"/><a:alphaOff val="10000"/></a:schemeClr></a:gs>`+second;
    if (variant === "gradient-one") stops=first;
    if (variant === "gradient-many") stops=first.repeat(257);
    if (variant === "gradient-max") stops=first.repeat(256);
    if (variant === "gradient-equal") stops=first+second.replace('pos="100000"','pos="0"');
    if (variant === "gradient-bad-pos") stops=first.replace('pos="0"','pos="-1"')+second;
    if (variant === "gradient-no-pos") stops=first.replace(' pos="0"','')+second;
    if (variant === "gradient-color-transform") stops=first.replace('<a:alpha val="50000"/>','<a:hueOff val="50000"/>')+second;
    const linear=variant === "gradient-defaults" ? '<a:lin/>' : variant === "gradient-path" ? '<a:path path="circle"/>' : variant === "gradient-bad-angle" ? '<a:lin ang="21600000"/>' : variant === "gradient-bad-scaled" ? '<a:lin scaled="maybe"/>' : '<a:lin ang="2700000" scaled="1"/>';
    const fill = `<a:gradFill${variant === "gradient-flip" ? ' flip="x"' : ''}${variant === "gradient-no-rotate" ? ' rotWithShape="0"' : ''}><a:gsLst>${stops}</a:gsLst>${linear}${variant === "gradient-tile" ? '<a:tileRect/>' : ''}${variant === "gradient-duplicate-linear" ? '<a:lin/>' : ''}</a:gradFill>${variant === "gradient-conflict" ? '<a:solidFill><a:srgbClr val="000000"/></a:solidFill>' : ''}`;
    const bg = `<p:bg><p:bgPr>${fill}</p:bgPr></p:bg>`;
    f.parts["deck/master.xml"] = f.parts["deck/master.xml"].replace('<p:cSld>', '<p:cSld>'+bg);
    f.parts["deck/layout.xml"] = f.parts["deck/layout.xml"].replace('<p:cSld>', '<p:cSld>'+bg);
    f.parts["deck/z.xml"] = f.parts["deck/z.xml"].replace('<p:sld ', `<p:sld xmlns:a="${f.A}" `).replace('<p:cSld name="First"/>', '<p:cSld name="First">'+bg+'</p:cSld>');
  }
  f.parts["deck/a.xml"] = f.parts["deck/a.xml"].replace('<p:sld ', '<p:sld showMasterSp="0" ');
  f.parts["deck/_rels/master.xml.rels"] = relationships([f.rel("layout","slideLayout","layout.xml"), f.rel("theme","theme","theme.xml"), f.rel("image","image","../media/image.png")]);
  f.parts["deck/_rels/layout.xml.rels"] = relationships([f.rel("master","slideMaster","master.xml"), f.rel("image","image","../media/image.png"), ...(["gradient-theme","gradient-theme-missing","theme-text","theme-missing","theme-map","theme-paint","paint-unsupported","paint-missing","line-unsupported","backgrounds","background-unsupported","background-matrix","background-matrix-missing","background-matrix-alpha","background-matrix-zero","background-matrix-gradient","background-matrix-unbound"].includes(variant) ? [f.rel("override","themeOverride","override.xml")] : [])]);
  f.parts["media/image.png"] = png;
  if (variant === 'decorative') {
    const extension = value => `<a:extLst><a:ext uri="{C183D7F6-B498-43B3-948B-1728B52AA6E4}"><d:decorative xmlns:d="http://schemas.microsoft.com/office/drawing/2017/decorative" val="${value}"/></a:ext></a:extLst>`;
    f.parts['deck/master.xml'] = f.parts['deck/master.xml'].replace('<p:cNvPr id="3" name="Picture 3"/>', `<p:cNvPr id="3" name="Picture 3">${extension('1')}</p:cNvPr>`);
    f.parts['deck/layout.xml'] = f.parts['deck/layout.xml'].replace('<p:cNvPr id="4" name="Picture 4"/>', `<p:cNvPr id="4" name="Picture 4">${extension('0')}</p:cNvPr>`);
  }
  return pack(f.parts,contentTypes(undefined,[...Object.entries(f.types).map(([PartName,type])=>({PartName:`/${PartName}`,ContentType:`application/vnd.openxmlformats-officedocument.${type}+xml`})),{PartName:"/media/image.png",ContentType:"image/png"}]));
}

test("opt-in shared artwork maps picture ownership, crop, asset reuse and both visibility switches", async () => {
  for (const strict of [false,true]) {
    const bytes = fixture(strict), session = await importPptx(bytes,{projection:"content",masterArtwork:"pictures-v1"}), { document, report } = session;
    const master = document.masters[0], layout = document.layouts.find(l => l.masterId === master.id);
    assert.equal(master.elements.length,1); assert.equal(layout.elements.length,1);
    assert.equal(master.elements[0].type,"image"); assert.deepEqual(master.elements[0].box,{x:1,y:2,width:10,height:20});
    assert.deepEqual(master.elements[0].crop,{left:0.1,top:0,right:0,bottom:0});
    assert.equal(master.elements[0].assetId,layout.elements[0].assetId); assert.equal(document.assets.length,1);
    assert.equal(layout.showMasterShapes,false); assert.equal(document.slides[0].showInheritedShapes,undefined); assert.equal(document.slides[1].showInheritedShapes,false);
    assert.deepEqual(document.slides.map(s => s.layoutId),[layout.id,layout.id]);
    assert.ok(report.findings.some(f => f.feature === "master-artwork.unmapped-property" && f.source.path.includes("sp[1]")));
    assert.equal(report.identities.filter(i=>i.kind==="master").length,1);
    const ready=await session.prepareSave({mode:"no-op"}); assert.equal(ready.status,"ready"); assert.deepEqual((await session.save(ready.plan)).bytes,new Uint8Array(bytes));
    assert.equal((await session.prepareSave({mode:"edit"})).status,"blocked");
    document.replaceSlideMaster(master.id,{name:master.name,elements:[{...master.elements[0],box:{...master.elements[0].box,x:9}}]});
    assert.equal((await session.prepareSave({mode:"no-op"})).status,"blocked"); session.dispose();
  }
});

test("artwork-v10 maps only explicit supported slide pattern fills in both namespaces",async()=>{
 for(const strict of [false,true]){
  for(const [kind,expected] of [["horz","horizontal"],["vert","vertical"],["cross","cross"],["dnDiag","downwardDiagonal"],["upDiag","upwardDiagonal"],["diagCross","diagonalCross"],...[5,10,20,25,30,40,50,60,70,75,80,90].map(n=>[`pct${n}`,`percent${n}`])]) {const s=await importPptx(fixture(strict,false,`slide-pattern-${kind}`),{projection:"content",masterArtwork:"artwork-v10"});assert.deepEqual(s.document.slides[0].elements[0].style.fill,{type:"pattern",pattern:expected,foreground:"FF0000",background:"0000FF",foregroundOpacity:.5,backgroundOpacity:1});s.dispose();}
  for(const kind of ["unknown","missing","no-preset","unresolved","transform"]) {const s=await importPptx(fixture(strict,false,`slide-pattern-${kind}`),{projection:"content",masterArtwork:"artwork-v10"});assert.equal(s.document.slides[0].elements.length,0);assert.ok(s.report.findings.some(f=>f.tier==="unsupported"&&f.reason.includes("Shape pattern is outside")));assert.ok(!s.report.findings.some(f=>f.feature==="drawing.pattern-dependencies"));s.dispose();}
  const transparent=await importPptx(fixture(strict,false,"slide-pattern-zero"),{projection:"content",masterArtwork:"artwork-v10"});assert.equal(transparent.document.slides[0].elements[0].style.fill.foregroundOpacity,0);transparent.dispose();
 }
});

test("artwork-v11 maps explicit outer shadows in both namespaces and reports unsupported effects",async()=>{
 for(const strict of [false,true]){
  const bytes=fixture(strict,false,"slide-shadow-valid"),session=await importPptx(bytes,{projection:"content",masterArtwork:"artwork-v11"});assert.deepEqual(session.document.slides[0].elements[0].style.shadow,{color:"000000",opacity:.5,blur:6,distance:8,direction:90});assert.ok(session.report.findings.some(f=>f.feature==="drawing.shadow-dependencies"));const ready=await session.prepareSave({mode:"no-op"});assert.deepEqual((await session.save(ready.plan)).bytes,new Uint8Array(bytes));session.dispose();
  const fallback=await importPptx(fixture(strict,false,"slide-shadow-no-blur"),{projection:"content",masterArtwork:"artwork-v11"});assert.equal(fallback.document.slides[0].elements[0].style.shadow.blur,0);assert.ok(fallback.report.findings.some(f=>f.reason.includes("Missing shadow blur radius")));fallback.dispose();
  for(const kind of ["scale","rotation","missing","color","multiple"]){const omitted=await importPptx(fixture(strict,false,`slide-shadow-${kind}`),{projection:"content",masterArtwork:"artwork-v11"});assert.equal(omitted.document.slides[0].elements[0].style.shadow,undefined);assert.ok(omitted.report.findings.some(f=>f.tier==="unsupported"&&f.source.path?.includes("effectLst")));omitted.dispose();}
 }
});

test("artwork-v9 maps slide gradient shapes in Strict and Transitional packages",async()=>{
 for(const strict of [false,true]){
  const bytes=fixture(strict,false,"slide-gradient"),session=await importPptx(bytes,{projection:"content",masterArtwork:"artwork-v9"});
  const shape=session.document.slides[0].elements[0];assert.equal(shape.type,"shape");assert.equal(shape.style.fill.angle,90);assert.equal(shape.style.fill.stops[0].color,"FF0000");session.dispose();
  const omitted=await importPptx(fixture(strict,false,"slide-gradient-flip"),{projection:"content",masterArtwork:"artwork-v9"});assert.equal(omitted.document.slides[0].elements.length,0);assert.ok(omitted.report.findings.some(f=>f.tier==="unsupported"&&f.reason.includes("Shape gradient is outside")));omitted.dispose();
 }
});

test("artwork-v8 imports direct and grouped shared shape linear gradients without changing earlier profiles",async()=>{
 for(const strict of [false,true]){
  const bytes=fixture(strict,false,"shape-gradient"),old=await importPptx(bytes,{projection:"content",masterArtwork:"artwork-v7"}),current=await importPptx(bytes,{projection:"content",masterArtwork:"artwork-v8"});
  assert.deepEqual(old.document.masters[0].elements.map(e=>e.type),["image"]);
  const shapes=[current.document.masters[0].elements[0],current.document.masters[0].elements[1].children[0],current.document.layouts.find(l=>l.masterId).elements[0]];
  const dependencies=current.report.findings.filter(f=>f.feature==="master-artwork.gradient-dependencies");assert.equal(dependencies.length,3);assert.ok(dependencies.every(f=>f.resolvedFrom?.some(x=>x.path.includes("gs[1]"))));
  for(const shape of shapes) assert.deepEqual(shape.style.fill,{type:"linearGradient",angle:45,scaled:true,stops:[{position:0,color:"FF0000",opacity:.5},{position:1,color:"0000FF",opacity:1}]});
  const ready=await current.prepareSave({mode:"no-op"});assert.deepEqual((await current.save(ready.plan)).bytes,new Uint8Array(bytes));assert.equal((await current.prepareSave({mode:"edit"})).status,"blocked");old.dispose();current.dispose();
  const unsupported=await importPptx(fixture(strict,false,"shape-gradient-unsupported"),{projection:"content",masterArtwork:"artwork-v8"});assert.deepEqual(unsupported.document.masters[0].elements.map(e=>e.type),["image"]);assert.ok(unsupported.report.findings.some(f=>f.feature==="master-artwork.shape" && f.tier==="unsupported"));unsupported.dispose();
  await assert.rejects(importPptx(bytes,{projection:"content",masterArtwork:"artwork-v8",sessionLimits:{maxTextStyleDependencies:1}}),{code:"import-limit-exceeded"});
 }
});

test("artwork-v7 projects bounded direct gradients and keeps older profiles unchanged", async () => {
 for (const strict of [false, true]) {
  const bytes=fixture(strict,false,"gradient-direct"), old=await importPptx(bytes,{projection:"content",masterArtwork:"artwork-v6"}), current=await importPptx(bytes,{projection:"content",masterArtwork:"artwork-v7"});
  assert.equal(old.document.masters[0].background,undefined);
  const expected={type:"linearGradient",angle:45,scaled:true,stops:[{position:0,color:"FF0000",opacity:.5},{position:1,color:"0000FF",opacity:1}]};
  assert.deepEqual(current.document.masters[0].background,expected); assert.deepEqual(current.document.layouts.find(l=>l.masterId).background,expected); assert.deepEqual(current.document.slides[0].background,expected);
  assert.deepEqual(normalizePresentation(current.document).slides.map(s=>s.backgroundSource),["slide","layout"]);
  const ready=await current.prepareSave({mode:"no-op"}); assert.deepEqual((await current.save(ready.plan)).bytes,new Uint8Array(bytes));old.dispose();current.dispose();
  for (const variant of ["gradient-unsorted","gradient-flip","gradient-tile","gradient-one","gradient-many","gradient-path","gradient-no-rotate","gradient-color-transform"]) { const session=await importPptx(fixture(strict,false,variant),{projection:"content",masterArtwork:"artwork-v7"}); assert.equal(session.document.masters[0].background,undefined);assert.ok(session.report.findings.some(f=>f.feature==="artwork.background" && f.tier==="unsupported"));session.dispose(); }
  const defaults=await importPptx(fixture(strict,false,"gradient-defaults"),{projection:"content",masterArtwork:"artwork-v7"});assert.equal(defaults.document.masters[0].background.angle,0);assert.equal(defaults.document.masters[0].background.scaled,false);assert.ok(defaults.report.findings.some(f=>f.feature==="artwork.background.unmapped-property" && f.reason.includes("inherited gradient attributes")));defaults.dispose();
  const max=await importPptx(fixture(strict,false,"gradient-max"),{projection:"content",masterArtwork:"artwork-v7"});assert.equal(max.document.masters[0].background.stops.length,256);max.dispose();
  const equal=await importPptx(fixture(strict,false,"gradient-equal"),{projection:"content",masterArtwork:"artwork-v7"});assert.deepEqual(equal.document.masters[0].background.stops.map(s=>s.position),[0,0]);equal.dispose();
  for(const variant of ["gradient-bad-pos","gradient-no-pos","gradient-bad-angle","gradient-bad-scaled","gradient-duplicate-linear","gradient-conflict"]) await assert.rejects(importPptx(fixture(strict,false,variant),{projection:"content",masterArtwork:"artwork-v7"}),{code:"import-invalid-package"},variant);
  await assert.rejects(importPptx(bytes,{projection:"content",masterArtwork:"artwork-v7",sessionLimits:{maxTextStyleDependencies:1}}),{code:"import-limit-exceeded"});
 }
});

test("gradient stops resolve theme and ordered alpha in each background owner scope",async()=>{
 for(const strict of [false,true]){
  const session=await importPptx(fixture(strict,false,"gradient-theme"),{projection:"content",masterArtwork:"artwork-v7"});
  const colors=[session.document.masters[0].background,session.document.layouts.find(l=>l.masterId).background,session.document.slides[0].background].map(p=>p.stops[0]);
  assert.deepEqual(colors,["123456","AA0000","FFFFFF"].map(color=>({position:0,color,opacity:.5})));
  const normal=normalizePresentation(session.document);assert.deepEqual(normal.slides.map(s=>s.background.stops[0].color),["FFFFFF","AA0000"]);
  for (const [part,theme] of [["/deck/master.xml","/deck/theme.xml"],["/deck/layout.xml","/deck/override.xml"],["/deck/z.xml","/deck/slide-override.xml"]]) {const trace=session.report.findings.find(f=>f.feature==="artwork.background"&&f.source.part===part).resolvedFrom;assert.ok(Object.isFrozen(trace));assert.ok(trace.some(x=>x.part===part&&x.path.includes("gs[1]")));assert.ok(trace.some(x=>x.part===part&&x.path.includes("gs[2]")));assert.ok(trace.some(x=>x.part===theme));}
  session.dispose();
  const missing=await importPptx(fixture(strict,false,"gradient-theme-missing"),{projection:"content",masterArtwork:"artwork-v7"});
  assert.equal(missing.document.masters[0].background,undefined);assert.ok(missing.report.findings.some(f=>f.feature==="artwork.background" && f.tier==="unsupported"));missing.dispose();
 }
});

test("artwork-v6 resolves background matrix placeholder colors per owner and blocks missing entries",async()=>{
 for(const strict of [false,true]) {
  const bytes=fixture(strict,false,"background-matrix"),prior=await importPptx(bytes,{projection:"content",masterArtwork:"artwork-v5"}),s=await importPptx(bytes,{projection:"content",masterArtwork:"artwork-v6"});
  assert.equal(prior.document.masters[0].background,undefined);
  assert.deepEqual(s.document.masters[0].background,{type:"solid",color:"123456",opacity:.5});
  assert.deepEqual(s.document.layouts.find(l=>l.masterId).background,{type:"solid",color:"AA0000",opacity:.5});
  assert.deepEqual(s.document.slides[0].background,{type:"solid",color:"FFFFFF",opacity:.5});
  assert.equal(s.report.findings.filter(x=>x.feature==="artwork.background" && x.tier==="fallback").length,3);
  const trace=s.report.findings.find(x=>x.feature==="artwork.background" && x.source.part==="/deck/layout.xml").resolvedFrom;assert.ok(trace.some(x=>x.part==="/deck/theme.xml" && x.path.includes("bgFillStyleLst")));assert.ok(trace.some(x=>x.part==="/deck/override.xml"));assert.ok(Object.isFrozen(trace));
  const ready=await s.prepareSave({mode:"no-op"});assert.deepEqual((await s.save(ready.plan)).bytes,new Uint8Array(bytes));prior.dispose();s.dispose();
  const missing=await importPptx(fixture(strict,false,"background-matrix-missing"),{projection:"content",masterArtwork:"artwork-v6"});
  assert.equal(missing.document.masters[0].background,undefined);assert.ok(missing.report.findings.some(x=>x.feature==="artwork.background" && x.tier==="unsupported"));missing.dispose();
  const alpha=await importPptx(fixture(strict,false,"background-matrix-alpha"),{projection:"content",masterArtwork:"artwork-v6"});assert.equal(alpha.document.masters[0].background.opacity,.2);alpha.dispose();
  for(const variant of ["background-matrix-zero","background-matrix-gradient","background-matrix-unbound"]) {const unsupported=await importPptx(fixture(strict,false,variant),{projection:"content",masterArtwork:"artwork-v6"});assert.equal(unsupported.document.masters[0].background,undefined);assert.ok(unsupported.report.findings.some(x=>x.feature==="artwork.background" && x.tier==="unsupported"));unsupported.dispose();}
  await assert.rejects(importPptx(bytes,{projection:"content",masterArtwork:"artwork-v6",sessionLimits:{maxTextStyleDependencies:1}}),{code:"import-limit-exceeded"});
 }
});

test("artwork-v5 maps explicit backgrounds in owner scope and preserves slide override precedence",async()=>{
 for(const strict of [false,true]) {
  const bytes=fixture(strict,false,"backgrounds"),prior=await importPptx(bytes,{projection:"content",masterArtwork:"artwork-v4"}),s=await importPptx(bytes,{projection:"content",masterArtwork:"artwork-v5"});
  assert.equal(prior.document.masters[0].background,undefined);assert.equal(prior.document.slides[0].background,undefined);
  assert.equal(s.document.masters[0].background.color,"123456");assert.equal(s.document.layouts.find(l=>l.masterId).background.color,"AA0000");assert.equal(s.document.slides[0].background.color,"FFFFFF");assert.equal(s.document.slides[1].background,undefined);
  const normal=normalizePresentation(s.document);assert.deepEqual(normal.slides.map(x=>x.background.color),["FFFFFF","AA0000"]);assert.deepEqual(normal.slides.map(x=>x.backgroundSource),["slide","layout"]);
  assert.equal(s.report.findings.filter(x=>x.feature==="artwork.background" && x.tier==="fallback").length,3);
  assert.ok(!s.report.findings.some(x=>x.feature==="drawing.unmapped-property" && x.source.part==="/deck/z.xml" && x.source.path?.includes("/p:bg[1]")));
  const ready=await s.prepareSave({mode:"no-op"});assert.deepEqual((await s.save(ready.plan)).bytes,new Uint8Array(bytes));
  s.document.slides[0].setBackground({type:"solid",color:"000000"});assert.equal((await s.prepareSave({mode:"no-op"})).status,"blocked");prior.dispose();s.dispose();
  const unsupported=await importPptx(fixture(strict,false,"background-unsupported"),{projection:"content",masterArtwork:"artwork-v5"});
  assert.equal(unsupported.document.slides[0].background,undefined);assert.ok(unsupported.report.findings.some(x=>x.feature==="artwork.background" && x.tier==="unsupported" && x.source.part==="/deck/z.xml"));unsupported.dispose();
 }
});

test("artwork-v4 resolves scoped shape fill and line theme paints without silent transform fallback",async()=>{
 for(const strict of [false,true]) {
  const bytes=fixture(strict,false,"theme-paint"),prior=await importPptx(bytes,{projection:"content",masterArtwork:"artwork-v3"}),session=await importPptx(bytes,{projection:"content",masterArtwork:"artwork-v4"});
  assert.deepEqual(prior.document.masters[0].elements.map(e=>e.type),["image"]);
  const master=session.document.masters[0].elements[0],layout=session.document.layouts.find(l=>l.masterId).elements[0];
  assert.deepEqual(master.style.fill,{type:"solid",color:"123456",opacity:.5});assert.equal(master.style.stroke.paint.color,"123456");
  assert.deepEqual(layout.style.fill,{type:"solid",color:"AA0000",opacity:.5});assert.equal(layout.style.stroke.paint.color,"AA0000");
  const ready=await session.prepareSave({mode:"no-op"});assert.deepEqual((await session.save(ready.plan)).bytes,new Uint8Array(bytes));prior.dispose();session.dispose();
  await assert.rejects(importPptx(bytes,{projection:"content",masterArtwork:"artwork-v4",sessionLimits:{maxTextStyleDependencies:1}}),{code:"import-limit-exceeded"});
  const bad=await importPptx(fixture(strict,false,"paint-unsupported"),{projection:"content",masterArtwork:"artwork-v4"});
  assert.deepEqual(bad.document.masters[0].elements.map(e=>e.type),["image"]);assert.ok(bad.report.findings.some(f=>f.feature==="master-artwork.unmapped-property" && f.reason.includes("Unsupported color transform")));bad.dispose();
  for (const variant of ["paint-missing","line-unsupported"]) {const omitted=await importPptx(fixture(strict,false,variant),{projection:"content",masterArtwork:"artwork-v4"});assert.deepEqual(omitted.document.masters[0].elements.map(e=>e.type),["image"]);assert.ok(omitted.report.findings.some(f=>f.feature==="master-artwork.unmapped-property" && (variant === "paint-missing" ? f.reason.includes("Theme slot") : f.reason.includes("Unsupported color transform"))));omitted.dispose();}
  const rgb=await importPptx(fixture(strict,false,"rgb-transform"),{projection:"content",masterArtwork:"artwork-v4"});
  assert.deepEqual(rgb.document.masters[0].elements[0].style.fill,{type:"solid",color:"000080",opacity:.5});rgb.dispose();
 }
});

test("artwork-v3 resolves text themes in separate master and layout scopes",async()=>{
 for(const strict of [false,true]) {
  const bytes=fixture(strict,false,"theme-text"),old=await importPptx(bytes,{projection:"content",masterArtwork:"artwork-v2"}),session=await importPptx(bytes,{projection:"content",masterArtwork:"artwork-v3"});
  const master=session.document.masters[0].elements[0],layout=session.document.layouts.find(l=>l.masterId).elements[0];
  assert.equal(master.text.content[0].runs[0].style.color,"123456");assert.equal(master.text.content[0].runs[0].style.fontFamily,"Base Font");
  assert.equal(layout.text.content[0].runs[0].style.color,"AA0000");assert.equal(layout.text.content[0].runs[0].style.fontFamily,"Layout Font");
  assert.equal(old.document.masters[0].elements[0].text.content[0].runs[0].style.color,undefined);
  assert.ok(session.report.textStyleInheritance.properties.some(p=>p.property==="run.color" && p.target.part==="/deck/layout.xml" && p.resolvedFrom?.some(s=>s.part==="/deck/override.xml")));
  const ready=await session.prepareSave({mode:"no-op"});assert.deepEqual((await session.save(ready.plan)).bytes,new Uint8Array(bytes));old.dispose();session.dispose();
  const missing=await importPptx(fixture(strict,false,"theme-missing"),{projection:"content",masterArtwork:"artwork-v3"});
  assert.equal(missing.document.masters[0].elements[0].text.content[0].runs[0].style.color,undefined);assert.ok(missing.report.findings.some(f=>f.feature==="master-artwork.unmapped-property" && f.reason.includes("Theme slot")));missing.dispose();
  const mapped=await importPptx(fixture(strict,false,"theme-map"),{projection:"content",masterArtwork:"artwork-v3"});
  assert.equal(mapped.document.layouts.find(l=>l.masterId).elements[0].text.content[0].runs[0].style.color,"00CC00");assert.equal(mapped.document.masters[0].elements[0].text.content[0].runs[0].style.color,"123456");mapped.dispose();
  await assert.rejects(importPptx(bytes,{projection:"content",masterArtwork:"artwork-v3",sessionLimits:{maxTextStyleProperties:1}}),{code:"import-limit-exceeded"});
 }
});

test("artwork-v2 projects direct text in shared shapes while reporting unsupported fields",async()=>{
 for(const strict of [false,true]) {
  const bytes=fixture(strict,false,"rich-text"), prior=await importPptx(bytes,{projection:"content",masterArtwork:"artwork-v1"}), session=await importPptx(bytes,{projection:"content",masterArtwork:"artwork-v2"});
  assert.deepEqual(prior.document.masters[0].elements.map(e=>e.type),["image"]);
  const [shape,group,image]=session.document.masters[0].elements;assert.deepEqual([shape.type,group.type,image.type],["shape","group","image"]);
  assert.deepEqual(shape.text.content[0].runs.map(r=>r.text),["你好 & test","\n","second"]);assert.equal(shape.text.content[0].runs[0].style.fontSize,24);assert.equal(shape.text.content[0].runs[0].style.bold,true);assert.equal(shape.text.frame.margin.left,1);
  assert.equal(group.children[0].text.content[0].runs[0].text,"你好 & test");assert.equal(shape.text.content[1].runs[0].text,"");
  assert.equal(session.document.layouts.find(l=>l.masterId).elements[0].text.content[0].runs[0].text,"Layout text");
  assert.equal(session.report.findings.filter(f=>f.feature==="master-artwork.shape-text").length,3);
  assert.ok(session.report.findings.some(f=>f.feature==="master-artwork.unmapped-property" && f.source.path?.includes("fld[1]")));
  assert.ok(session.report.findings.some(f=>f.feature==="master-artwork.unmapped-property" && f.source.path?.includes("hlinkClick[1]")));assert.equal(shape.text.content[0].runs[0].action,undefined);
  const ready=await session.prepareSave({mode:"no-op"});assert.deepEqual((await session.save(ready.plan)).bytes,new Uint8Array(bytes));
  assert.equal((await session.prepareSave({mode:"edit"})).status,"blocked");prior.dispose();session.dispose();
 }
});

test("artwork profile adds eligible direct and grouped shapes in source order",async()=>{
 for(const strict of [false,true]) {
  const bytes=fixture(strict,false,"shapes"), pictures=await importPptx(bytes,{projection:"content",masterArtwork:"pictures-v1"}), session=await importPptx(bytes,{projection:"content",masterArtwork:"artwork-v1"});
  assert.deepEqual(pictures.document.masters[0].elements.map(e=>e.type),["image"]);
  const elements=session.document.masters[0].elements;assert.deepEqual(elements.map(e=>e.type),["shape","group","image"]);
  assert.equal(elements[0].shape,"rect");assert.equal(elements[0].style.fill.color,"0000FF");assert.equal(elements[0].style.stroke.paint.type,"none");
  assert.equal(elements[1].children[0].type,"shape");assert.equal(elements[1].children[0].box.x,-2);assert.equal(elements[1].children[0].box.y,-1);
  assert.equal(session.report.findings.filter(f=>f.feature==="master-artwork.shape" && f.tier==="fallback").length,2);
  const ready=await session.prepareSave({mode:"no-op"});assert.deepEqual((await session.save(ready.plan)).bytes,new Uint8Array(bytes));
  pictures.dispose();session.dispose();
 }
 for(const variant of ["shape-text","shape-theme","shape-adjust","shape-no-line"]) {
  const session=await importPptx(fixture(false,false,variant),{projection:"content",masterArtwork:"artwork-v1"});
  assert.deepEqual(session.document.masters[0].elements.map(e=>e.type),["image"]);assert.ok(session.report.findings.some(f=>f.feature==="master-artwork.shape" && f.tier==="unsupported"));session.dispose();
 }
});

test("nested master groups retain local rebasing, transforms, shared asset and source identities",async()=>{
 for(const strict of [false,true]) {
  const session=await importPptx(fixture(strict,false,"nested"),{projection:"content",masterArtwork:"pictures-v1"});
  const outer=session.document.masters[0].elements[0],inner=outer.children[0],pic=inner.children[0];
  assert.equal(outer.type,"group");assert.deepEqual(outer.box,{x:10,y:20,width:40,height:30});assert.deepEqual(outer.coordinateSize,{width:100,height:20});
  assert.equal(outer.transform.rotation,1);assert.equal(outer.transform.flipH,true);
  assert.equal(inner.type,"group");assert.equal(inner.box.x,2);assert.equal(inner.box.y,2);assert.deepEqual(inner.coordinateSize,{width:20,height:20});
  assert.equal(pic.type,"image");assert.equal(pic.box.x,0);assert.equal(pic.box.y,1);assert.equal(pic.assetId,session.document.layouts.find(l=>l.masterId).elements[0].assetId);
  assert.equal(session.report.identities.filter(i=>i.kind==="element" && i.source.part==="/deck/master.xml").length,3);
  session.dispose();
 }
 const rejected=await importPptx(fixture(false,false,"zero-group"),{projection:"content",masterArtwork:"pictures-v1"});
 assert.deepEqual(rejected.document.masters[0].elements,[]);assert.ok(rejected.report.findings.some(f=>f.feature==="master-artwork.group" && f.reason.includes("Zero group")));rejected.dispose();
 for (const variant of ["duplicate","bad-id"]) await assert.rejects(importPptx(fixture(false,false,variant),{projection:"content",masterArtwork:"pictures-v1"}),{code:"import-invalid-package"});
});

test("group depth is bounded before a Core definition is constructed",async()=>{
 const atLimit=await importPptx(fixture(false,false,"depth48"),{projection:"content",masterArtwork:"pictures-v1"});
 let item=atLimit.document.masters[0].elements[0],depth=0;while(item.type==="group"){depth++;item=item.children[0]};assert.equal(depth,48);assert.equal(item.type,"image");atLimit.dispose();
 const over=await importPptx(fixture(false,false,"depth49"),{projection:"content",masterArtwork:"pictures-v1"});
 assert.deepEqual(over.document.masters[0].elements,[]);assert.ok(over.report.findings.some(f=>f.reason.includes("48-level")));over.dispose();
});

test("placeholder and grouped pictures are reported and never flattened into ordinary master artwork", async () => {
  for (const variant of ["placeholder","group"]) {
    const session=await importPptx(fixture(false,false,variant),{projection:"content",masterArtwork:"pictures-v1"});
    assert.deepEqual(session.document.masters[0].elements,[]);
    assert.ok(session.report.findings.some(f=>f.source.part==="/deck/master.xml" && f.tier==="unsupported" && (variant === "placeholder" ? f.feature==="master-artwork.picture" : f.source.path?.includes("grpSp[1]"))));
    session.dispose();
  }
});

test("unlisted layout paths are not attached, and profiles cannot imply preservation support", async () => {
  const bytes=fixture(false,true), session=await importPptx(bytes,{projection:"content",masterArtwork:"pictures-v1"});
  assert.ok(session.document.slides.every(s=>s.layoutId===undefined)); session.dispose();
  for (const options of [{projection:"structure",masterArtwork:"pictures-v1"},{projection:"content",masterArtwork:"pictures-v1",editSaveProfile:"text-runs-v1"},{projection:"content",masterArtwork:"pictures-v1",chartEditSaveProfile:"existing-values-v1"},{projection:"content",masterArtwork:"artwork-v1",editSaveProfile:"text-runs-v2"},{projection:"content",masterArtwork:"artwork-v1",chartEditSaveProfile:"existing-values-v1"},{projection:"content",masterArtwork:"artwork-v2",editSaveProfile:"text-runs-v3"},{projection:"content",masterArtwork:"artwork-v2",chartEditSaveProfile:"existing-values-v1"},{projection:"content",masterArtwork:"artwork-v3",editSaveProfile:"text-runs-v1"},{projection:"content",masterArtwork:"artwork-v3",chartEditSaveProfile:"existing-values-v1"},{projection:"content",masterArtwork:"artwork-v4",editSaveProfile:"text-runs-v1"},{projection:"content",masterArtwork:"artwork-v5",editSaveProfile:"text-runs-v1"},{projection:"content",masterArtwork:"artwork-v5",chartEditSaveProfile:"existing-values-v1"},{projection:"content",masterArtwork:"artwork-v6",editSaveProfile:"text-runs-v1"},{projection:"content",masterArtwork:"artwork-v6",chartEditSaveProfile:"existing-values-v1"},{projection:"content",masterArtwork:"invalid"}]) await assert.rejects(importPptx(bytes,options),TypeError);
});
