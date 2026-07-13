import type { TextStyle } from "../types/style.js";

export function assertTextStyle(style: TextStyle | undefined, context: string): void {
  if (style?.lineSpacing !== undefined && (!Number.isFinite(style.lineSpacing) || style.lineSpacing <= 0)) {
    throw new Error(`${context} lineSpacing must be a positive finite number.`);
  }

  const fontScale = style?.autoFit?.mode === "shrink" ? style.autoFit.fontScale : undefined;
  if (fontScale !== undefined && (!Number.isFinite(fontScale) || fontScale <= 0 || fontScale > 1)) {
    throw new Error(`${context} autoFit fontScale must be greater than zero and at most one.`);
  }
}
