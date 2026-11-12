import type { TextRunStyleInput, TextParagraphStyleInput, TextFrameStyleInput, TextStylePresetResetProperty } from "../types/style.js";
import { TEXT_FONT_SCRIPT_CODES } from "./font-scripts.js";

/** The independently resettable run declarations. */
export const TEXT_RUN_PROPERTIES = [
  "fontFamily", "eastAsianFontFamily", "complexScriptFontFamily", "scriptFontFamilies", "rightToLeft",
  "fontSize", "bold", "italic", "underline", "strike", "color", "colorOpacity", "language",
] as const satisfies readonly (keyof TextRunStyleInput)[];

export function isTextRunProperty(value: unknown): value is keyof TextRunStyleInput {
  return typeof value === "string" && (TEXT_RUN_PROPERTIES as readonly string[]).includes(value);
}

const PARAGRAPH_PROPERTIES = ["align", "direction", "indent", "hanging", "lineSpacing", "spaceBefore", "spaceAfter", "bullet", "tabStops"] as const satisfies readonly (keyof TextParagraphStyleInput)[];
const FRAME_PROPERTIES = ["margin", "verticalAlign", "wrap", "autoFit", "textDirection"] as const satisfies readonly (keyof TextFrameStyleInput)[];
export function isTextStylePresetResetProperty(value: unknown): value is TextStylePresetResetProperty {
  if (isTextRunProperty(value)) return true;
  if (typeof value !== "string") return false;
  if (value.startsWith("scriptFontFamilies.")) return (TEXT_FONT_SCRIPT_CODES as readonly string[]).includes(value.slice(19));
  if (["frame.margin.top", "frame.margin.right", "frame.margin.bottom", "frame.margin.left"].includes(value)) return true;
  if (value.startsWith("paragraph.")) return (PARAGRAPH_PROPERTIES as readonly string[]).includes(value.slice(10));
  if (value.startsWith("frame.")) return (FRAME_PROPERTIES as readonly string[]).includes(value.slice(6));
  return false;
}
