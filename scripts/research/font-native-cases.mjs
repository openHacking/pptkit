import { createPresentation, normalizePresentation, resolveTextFontSegments } from "../../packages/core/dist/index.js";
import { generatePptx } from "../../packages/pptx-exporter/dist/index.js";
import { importPptx } from "../../packages/pptx-parser/dist/index.js";
import { readGeneratedZip } from "../../tests/support/read-generated-zip.mjs";
import { zip } from "../../packages/pptx-parser/test/zip-fixture.mjs";

export const nativeFontFamilies = Object.freeze({ latin: "Arial", generic: "Arial Unicode MS", sc: "Hiragino Sans GB", tc: "Songti TC", ja: "Hiragino Mincho ProN", ko: "Apple SD Gothic Neo", ar: "Geeza Pro", he: "Arial Hebrew", greek: "Georgia" });
const F = nativeFontFamilies;
const theme = ["+mn-lt", "+mn-ea", "+mn-cs"];
export const nativeFontCases = [
  { id: "direct-sc", text: "骨直令返", language: "zh-CN", slots: [F.latin, F.sc, F.ar] },
  { id: "direct-arabic", text: "لا", language: "ar", slots: [F.latin, F.sc, F.ar] },
  { id: "direct-hebrew", text: "שָׁלוֹם", language: "he", slots: [F.latin, F.sc, F.he] },
  ...[["sc", "zh-CN", "骨直令返"], ["tc", "zh-TW", "骨直令返"], ["ja", "ja", "骨直令返"], ["ko", "ko", "骨直令返"], ["arabic", "ar", "لا"], ["hebrew", "he", "שָׁלוֹם"], ["greek", "el", "ΩΣ"]].map(([id, language, text]) => ({ id: `theme-explicit-${id}`, text, language, slots: theme })),
  ...[["sc", "zh-CN", "骨直令返"], ["tc", "zh-TW", "骨直令返"], ["arabic", "ar", "لا"], ["hebrew", "he", "שָׁלוֹם"]].map(([id, language, text]) => ({ id: `theme-empty-${id}`, text, language, slots: theme, empty: true })),
  { id: "latin-only-tc", text: "骨直令返", language: "zh-TW", slots: ["+mn-lt", null, null] },
  { id: "cross-slot-ea", text: "骨直令返", language: "zh-TW", slots: ["+mn-lt", "+mn-lt", "+mn-cs"] },
  { id: "local-ea-overrides", text: "骨直令返", language: "zh-TW", slots: [null, F.sc, null], inherited: true },
  { id: "local-cs-overrides", text: "שָׁלוֹם", language: "he", slots: [null, null, F.generic], inherited: true },
  { id: "bare-zh-empty", text: "骨直令返", language: "zh", slots: theme, empty: true },
  { id: "missing-glyph-latin", text: "骨直令返", language: "zh-TW", slots: [F.latin, F.latin, F.latin] },
];
// Keep the original 20 controls stable; these pairs isolate coverage from selection.
export const nativeFontPrecedenceCases = [
  { id: "ko-hangul-direct", text: "한글", language: "ko", slots: [F.latin, F.ko, F.ar] },
  { id: "ko-hangul-theme", text: "한글", language: "ko", slots: theme },
  { id: "ko-hangul-empty-theme", text: "한글", language: "ko", slots: theme, empty: true },
  { id: "ko-han-covered-map", text: "骨直令返", language: "ko", slots: theme, supplemental: { Hang: F.generic } },
  { id: "ko-hangul-covered-map", text: "한글", language: "ko", slots: theme, supplemental: { Hang: F.generic } },
  { id: "ko-kr-hangul-theme", text: "한글", language: "ko-KR", slots: theme },
  { id: "ko-kr-hangul-empty", text: "한글", language: "ko-KR", slots: theme, empty: true },
  { id: "ko-kr-han-direct", text: "骨直令返", language: "ko-KR", slots: [F.latin, F.ko, F.ar] },
  { id: "ko-kr-han-theme", text: "骨直令返", language: "ko-KR", slots: theme },
  { id: "ko-kr-han-alternate", text: "骨直令返", language: "ko-KR", slots: theme, supplemental: { Hang: F.sc } },
  ...[["tc", "zh-TW", "骨直令返"], ["arabic", "ar", "لا"], ["hebrew", "he", "שָׁלוֹם"]].flatMap(([id, language, text]) => [
    { id: `absent-${id}-slots`, text, language, slots: ["+mn-lt", null, null], themeSlots: [F.generic, F.generic, F.generic] },
    { id: `present-${id}-slots`, text, language, slots: theme, themeSlots: [F.generic, F.generic, F.generic] },
    { id: `empty-local-${id}`, text, language, slots: ["+mn-lt", "", ""], themeSlots: [F.generic, F.generic, F.generic] },
  ]),
  { id: "cross-ea-latin-covered", text: "骨直令返", language: "zh-TW", slots: ["+mn-lt", "+mn-lt", "+mn-cs"], themeSlots: [F.sc, F.generic, F.generic] },
  { id: "literal-ea-latin-covered", text: "骨直令返", language: "zh-TW", slots: [F.sc, F.sc, F.generic] },
  { id: "absent-tc-no-hant-map", text: "骨直令返", language: "zh-TW", slots: ["+mn-lt", null, null], themeSlots: [F.generic, F.generic, F.generic], supplemental: { Hant: null } },
  { id: "present-tc-no-hant-map", text: "骨直令返", language: "zh-TW", slots: theme, supplemental: { Hant: null } },
  { id: "present-tc-alternate-map", text: "骨直令返", language: "zh-TW", slots: theme, supplemental: { Hant: F.sc } },
  ...[["sc", "zh-CN"], ["tc", "zh-TW"], ["ja", "ja"]].map(([id, language]) => ({ id: `major-${id}-map`, text: "骨直令返", language, slots: ["+mj-lt", "+mj-ea", "+mj-cs"], collection: "major" })),
  { id: "bare-zh-explicit", text: "骨直令返", language: "zh", slots: theme },
  { id: "bare-zh-no-hans-map", text: "骨直令返", language: "zh", slots: theme, empty: true, supplemental: { Hans: null } },
  { id: "bare-zh-hans-alternate", text: "骨直令返", language: "zh", slots: theme, empty: true, supplemental: { Hans: F.tc } },
  { id: "zh-hans-explicit", text: "骨直令返", language: "zh-Hans", slots: theme, empty: true },
  { id: "zh-hant-explicit", text: "骨直令返", language: "zh-Hant", slots: theme, empty: true },
];
const xml = value => value.replaceAll("&", "&amp;").replaceAll('"', "&quot;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
const fontSlots = slots => slots.map((family, index) => family === null ? "" : `<a:${["latin", "ea", "cs"][index]} typeface="${xml(family)}"/>`).join("");
export async function buildNativeFontCase(control) {
  const p = createPresentation({ size: { width: 720, height: 260 }, theme: { fonts: { body: F.latin, heading: F.latin } } }), slide = p.addSlide();
  slide.addElement({ type: "text", name: "Label", box: { x: 20, y: 20, width: 680, height: 40 }, content: [{ runs: [{ text: control.id, style: { fontFamily: F.latin, fontSize: 16 } }] }] });
  slide.addElement({ type: "text", name: "Probe", box: { x: 20, y: 90, width: 680, height: 110 }, frame: { margin: 0 }, content: [{ runs: [{ text: control.text, style: { fontFamily: F.latin, fontSize: 48, language: control.language } }] }] });
  const parts = readGeneratedZip((await generatePptx(p)).bytes);
  const supplements = Object.entries({ Hans: F.sc, Hant: F.tc, Jpan: F.ja, Hang: F.ko, Arab: F.ar, Hebr: F.he, Grek: F.greek, ...control.supplemental }).filter(([, family]) => family !== null).map(([script, family]) => `<a:font script="${script}" typeface="${xml(family)}"/>`).join("");
  const collection = control.collection === "major" ? "majorFont" : "minorFont";
  parts.set("ppt/theme/theme1.xml", Buffer.from(parts.get("ppt/theme/theme1.xml").toString().replace(new RegExp(`<a:${collection}>[\\s\\S]*?</a:${collection}>`), `<a:${collection}>${fontSlots(control.themeSlots ?? [F.latin, control.empty ? "" : F.generic, control.empty ? "" : F.generic])}${supplements}</a:${collection}>`)));
  parts.set("ppt/slides/slide1.xml", Buffer.from(parts.get("ppt/slides/slide1.xml").toString().replace(/<p:sp>[\s\S]*?<\/p:sp>/g, shape => {
    if (!shape.includes('name="Probe"')) return shape;
    let result = shape.replace(/<a:rPr\b([^>]*)>[\s\S]*?<\/a:rPr>/, (all, attributes) => `<a:rPr${attributes}>${all.slice(all.indexOf(">") + 1, all.lastIndexOf("</a:rPr>")).replace(/<a:(?:latin|ea|cs)\b[^>]*\/>/g, "")}${fontSlots(control.slots)}</a:rPr>`);
    if (control.inherited) result = result.replace('<a:lstStyle/>', `<a:lstStyle><a:defPPr><a:defRPr>${fontSlots(theme)}</a:defRPr></a:defPPr></a:lstStyle>`);
    return result;
  })));
  const bytes = zip([...parts].map(([name, body]) => ({ name, body }))).bytes;
  const session = await importPptx(bytes, { projection: "content", textFontProfile: "script-families-v2" });
  try {
    const document = normalizePresentation(session.document), probe = document.slides[0].elements.find(e => e.name === "Probe"), run = probe.content[0].runs[0];
    return { bytes, projection: { text: run.text, style: run.style, selected: resolveTextFontSegments(run.text, run.style, document.theme), properties: session.report.textStyleInheritance.properties.filter(p => p.target.sourceObjectId === session.report.identities.find(i => i.id === probe.id)?.source.sourceObjectId) } };
  } finally { session.dispose(); }
}
