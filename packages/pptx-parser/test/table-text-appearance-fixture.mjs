import { tableStyleFixture } from "./table-style-fixture.mjs";

export const textColor = (value = "accent2", modifiers = "", kind = "schemeClr") => `<a:${kind} val="${value}">${modifiers}</a:${kind}>`;
export const textReference = (color = textColor(), index = "minor") => `<a:fontRef idx="${index}">${color}</a:fontRef>`;
export const fontCollection = (latin = "Arial", ea = "Arial", cs = "Arial", extra = "") => `<a:font><a:latin typeface="${latin}"/><a:ea typeface="${ea}"/><a:cs typeface="${cs}"/>${extra}</a:font>`;
export function textAppearanceFixture({ strict = false, text = textReference(), extraRegions = "", flags = "", paragraphs } = {}) {
  const f = tableStyleFixture({ strict, flags, rows: 2, columns: 2, rules: `<a:wholeTbl><a:tcTxStyle>${text}</a:tcTxStyle></a:wholeTbl>${extraRegions}` });
  f.parts["deck/theme.xml"] = f.parts["deck/theme.xml"].replaceAll('<a:latin typeface="Arial"/>', '<a:latin typeface="Arial"/><a:ea typeface="Arial"/><a:cs typeface="Arial"/>');
  if (paragraphs !== undefined) f.parts["deck/z.xml"] = f.parts["deck/z.xml"].replace(/<a:p>[\s\S]*?<\/a:p>/g, paragraphs);
  return f;
}
