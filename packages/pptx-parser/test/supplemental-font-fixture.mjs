import { scriptFontImportFixture, sourceFonts } from "./script-font-import-fixture.mjs";

export const supplementalFonts = '<a:font script="Hans" typeface="Simplified Source"/><a:font script="Hant" typeface="Traditional Source"/><a:font script="Jpan" typeface="Japanese Source"/><a:font script="Hang" typeface="Korean Source"/><a:font script="Arab" typeface="Arabic Source"/><a:font script="Hebr" typeface="Hebrew Source"/><a:font script="Grek" typeface="Greek Source"/>';
export function supplementalFontFixture({ strict = false, run = sourceFonts("+mn-lt", "+mn-ea", "+mn-cs"), inherited = false, entries = supplementalFonts, language = "zh-TW", emptySlots = false } = {}) {
  const f = scriptFontImportFixture({ strict, run, inherited });
  f.parts["deck/theme.xml"] = f.parts["deck/theme.xml"].replace('</a:minorFont>', `${entries}</a:minorFont>`);
  if (emptySlots) f.parts["deck/theme.xml"] = f.parts["deck/theme.xml"].replace('typeface="Body East"', 'typeface=""').replace('typeface="Body Complex"', 'typeface=""');
  for (const part of ["deck/z.xml", "deck/a.xml"]) f.parts[part] = f.parts[part].replaceAll('<a:rPr>', `<a:rPr lang="${language}">`).replaceAll('Latin中文العربية', 'A中عאΩ');
  return f;
}
