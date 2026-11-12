import { createPresentation } from "../../packages/core/dist/index.js";
import { generatePptx } from "../../packages/pptx-exporter/dist/index.js";
import { readGeneratedZip } from "../support/read-generated-zip.mjs";
import { zip } from "../../packages/pptx-parser/test/zip-fixture.mjs";
import { nativeFontFamilies as F } from "../../scripts/research/font-native-cases.mjs";

const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const xml = text => text.replaceAll("&", "&amp;").replaceAll('"', "&quot;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
export const tableFontMaps = {
  minor: { Latn: F.latin, Hans: F.sc, Hant: F.tc, Jpan: F.ja, Hang: F.ko, Arab: F.ar, Hebr: F.he, Grek: F.greek },
  major: { Latn: F.greek, Hans: F.sc, Hant: F.sc, Jpan: F.sc, Hang: F.ko, Arab: F.ar, Hebr: F.he, Grek: F.greek },
  collection: { Latn: F.greek, Hans: F.ja, Hant: F.ja, Jpan: F.ja, Hang: F.ko, Arab: F.ar, Hebr: F.he, Grek: F.greek },
  "collection-generic": { Latn: F.greek, Hans: F.ja, Hant: F.ja, Jpan: F.ja, Hang: F.ko, Arab: F.ar, Hebr: F.he, Grek: F.greek },
};
const scripts = [
  ["latin", "Abc", "en-US", "Latn"], ["sc", "骨直令返", "zh-CN", "Hans"], ["tc", "骨直令返", "zh-TW", "Hant"],
  ["ja", "骨直令返", "ja", "Jpan"], ["ko", "한글", "ko", "Hang"], ["ko-kr", "한글", "ko-KR", "Hang"],
  ["ar", "السَّلَامُ", "ar", "Arab"], ["he", "שָׁלוֹם", "he", "Hebr"], ["greek", "Ω", "el", "Grek"],
];
const ids = { "collection-generic": "{A70A0004-0000-4000-8000-000000000004}", minor: "{A70A0001-0000-4000-8000-000000000001}", major: "{A70A0002-0000-4000-8000-000000000002}", collection: "{A70A0003-0000-4000-8000-000000000003}" };
const rules = map => Object.entries(map).map(([script, family]) => `<a:font script="${script}" typeface="${xml(family)}"/>`).join("");
const collection = (map, latin = F.latin, generic = "") => `<a:latin typeface="${xml(latin)}"/><a:ea typeface="${xml(generic)}"/><a:cs typeface="${xml(generic)}"/>${rules(map)}`;

/** Named source-control assembler. It patches table references/themes before import; not a fresh-author round trip. */
export async function nativeTableFontBytes(variants = ["minor", "major", "collection"]) {
  if (variants.some(kind => !Object.hasOwn(ids, kind))) throw Error("Unknown named table-font control");
  const document = createPresentation({ size: { width: 720, height: 260 }, theme: { fonts: { body: F.latin, heading: F.latin } } }), controls = [];
  for (const kind of variants) for (const [sampleId, text, language, script] of scripts) {
    const id = `${kind}-${sampleId}`, slide = document.addSlide({ id });
    slide.addElement({ id: `label-${id}`, type: "text", name: "Label", box: { x: 20, y: 20, width: 680, height: 40 }, content: [{ runs: [{ text: id, style: { fontFamily: F.latin, fontSize: 16 } }] }] });
    slide.addElement({ id: `table-${id}`, type: "table", name: "ProbeTable", box: { x: 20, y: 90, width: 680, height: 110 }, columns: [680], rows: [{ height: 110, cells: [{ content: [{ runs: [{ text, style: { fontFamily: F.latin, fontSize: 48, language, color: "000000" } }] }] }] }] });
    controls.push({ id, kind, sampleId, text, language, script, styleId: ids[kind], targetFamily: tableFontMaps[kind][script], slideId: id, elementId: `table-${id}` });
  }
  const parts = readGeneratedZip((await generatePptx(document)).bytes);
  let theme = parts.get("ppt/theme/theme1.xml").toString();
  for (const kind of ["major", "minor"]) theme = theme.replace(new RegExp(`<a:${kind}Font>[\\s\\S]*?</a:${kind}Font>`), `<a:${kind}Font>${collection(tableFontMaps[kind], kind === "major" ? F.greek : F.latin)}</a:${kind}Font>`);
  parts.set("ppt/theme/theme1.xml", Buffer.from(theme));
  const styles = variants.map(kind => {
    const id = ids[kind];
    const font = kind.startsWith("collection") ? `<a:font>${collection(tableFontMaps[kind], F.greek, kind === "collection-generic" ? F.generic : "")}</a:font>` : `<a:fontRef idx="${kind}"><a:srgbClr val="000000"/></a:fontRef>`;
    return `<a:tblStyle styleId="${id}" styleName="PPTKit native ${kind}"><a:wholeTbl><a:tcTxStyle>${font}<a:srgbClr val="000000"/></a:tcTxStyle><a:tcStyle><a:fill><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:fill></a:tcStyle></a:wholeTbl></a:tblStyle>`;
  }).join("");
  parts.set("ppt/tableStyles.xml", Buffer.from(`<a:tblStyleLst xmlns:a="${A}" def="${ids[variants[0]]}">${styles}</a:tblStyleLst>`));
  for (const [index, control] of controls.entries()) {
    const file = `ppt/slides/slide${index + 1}.xml`;
    let replacements = 0;
    const slide = parts.get(file).toString().replace(/<a:tbl>[\s\S]*?<\/a:tbl>/, table => {
      replacements++;
      return table.replace(/<a:tblPr\b[^>]*>[\s\S]*?<\/a:tblPr>/, `<a:tblPr><a:tableStyleId>${control.styleId}</a:tableStyleId></a:tblPr>`)
        .replace(/<a:rPr\b[^>]*>[\s\S]*?<\/a:rPr>/g, `<a:rPr lang="${control.language}" sz="4800"/>`);
    });
    if (replacements !== 1) throw Error("Expected exactly one generated table per source control");
    parts.set(file, Buffer.from(slide));
  }
  return { bytes: zip([...parts].map(([name, body]) => ({ name, body }))).bytes, controls, maps: tableFontMaps };
}
