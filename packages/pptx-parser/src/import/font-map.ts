import { TEXT_FONT_SCRIPT_CODES, type TextFontScriptCode } from "@pptkit/core";
import { children, invalidStructure } from "../presentation/structure-xml.js";
import type { PptxSourceLocation } from "../types/source.js";
import type { XmlElement } from "../xml/read-xml.js";
import type { ContentReader } from "./content-reader.js";

export interface SourceFontEntry {
  source: PptxSourceLocation;
  status: "mapped" | "unsupported";
  value?: string;
  reason?: string;
  resolvedFrom?: readonly PptxSourceLocation[];
}
export type SourceFontMap = ReadonlyMap<TextFontScriptCode, SourceFontEntry>;
const admitted = new Set<string>(TEXT_FONT_SCRIPT_CODES);
const eastern = new Set<TextFontScriptCode>(["Hani", "Hans", "Hant", "Hira", "Kana", "Jpan", "Hang", "Kore", "Bopo", "Yiii"]);
const general = new Set<TextFontScriptCode>(["Zyyy", "Latn", "Grek", "Cyrl"]);
/** Native destination bucket, independent of current characters or language. */
export const sourceFontField = (code: TextFontScriptCode) => eastern.has(code) ? "run.eastAsianFontFamily" : general.has(code) ? "run.fontFamily" : "run.complexScriptFontFamily";

/** Read a collection once. Unknown identities stay reported; malformed identities block its map. */
export async function readSourceFontMap(node: XmlElement, c: { r: ContentReader; at: (node: XmlElement, attr?: string) => PptxSourceLocation }, checkpoint: () => Promise<void>): Promise<SourceFontMap> {
  const entries = new Map<TextFontScriptCode, SourceFontEntry>(), seen = new Set<string>();
  let invalid: SourceFontEntry | undefined;
  for (const font of children(node)) {
    await checkpoint();
    if (font.namespaceUri !== c.r.a || font.local !== "font") continue;
    c.r.known(font);
    const script = c.r.attr(font, "script"), family = c.r.attr(font, "typeface");
    if (!script) {
      const reason = "A supplemental font without a script identity blocks the collection's script map.";
      c.r.note(font, reason); invalid = { status: "unsupported", source: c.at(font), reason }; continue;
    }
    if (seen.has(script)) invalidStructure("Duplicate supplemental font script identities.", c.r.part);
    seen.add(script);
    if (!admitted.has(script)) { c.r.note(font, `Supplemental script ${JSON.stringify(script)} is outside the admitted Core script map and remains source-only.`); continue; }
    const source = c.at(font, "typeface"), resolvedFrom = [c.at(node), c.at(font, "script")];
    if (!family?.trim() || family.startsWith("+") || /[\u0000-\u001F\u007F]/u.test(family)) {
      const reason = "Supplemental typeface must be an explicit nonempty family without controls or theme recursion.";
      c.r.note(font, reason); entries.set(script as TextFontScriptCode, { source, status: "unsupported", reason, resolvedFrom });
    } else entries.set(script as TextFontScriptCode, { source, status: "mapped", value: family, resolvedFrom });
  }
  if (invalid) for (const code of TEXT_FONT_SCRIPT_CODES) entries.set(code, invalid);
  return entries;
}
