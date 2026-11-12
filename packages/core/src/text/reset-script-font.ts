import type { TextRunStyleInput } from "../types/style.js";
import { TEXT_FONT_SCRIPT_CODES, type TextFontScriptCode } from "./font-scripts.js";

/** Retain all other explicit clears when removing one key from a whole-map barrier. */
export function resetScriptFont(style: TextRunStyleInput | undefined, code: TextFontScriptCode): void {
  if (style?.scriptFontFamilies === undefined) return;
  const map = style.scriptFontFamilies === null
    ? Object.fromEntries(TEXT_FONT_SCRIPT_CODES.map(key => [key, null]))
    : { ...style.scriptFontFamilies };
  delete map[code];
  style.scriptFontFamilies = map;
}
