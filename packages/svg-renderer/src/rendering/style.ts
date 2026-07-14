import type {
  ColorValue,
  NormalizedPaint,
  NormalizedPresentationTheme,
  NormalizedStrokeStyle,
  NormalizedTransform,
  Box,
} from "@pptkit/core";

export function colorValue(value: ColorValue, theme: NormalizedPresentationTheme): string {
  const raw = typeof value === "string" ? value : theme.colors[value.theme];
  return raw.startsWith("#") ? raw : `#${raw}`;
}

export function paintAttributes(
  paint: NormalizedPaint,
  theme: NormalizedPresentationTheme,
  kind: "fill" | "stroke",
): string {
  if (paint.type === "none") return `${kind}="none"`;
  return `${kind}="${colorValue(paint.color, theme)}" ${kind}-opacity="${paint.opacity}"`;
}

export function dashArray(style: NormalizedStrokeStyle): string | undefined {
  const width = Math.max(style.width, 1);
  if (style.dash === "dash") return `${4 * width} ${3 * width}`;
  if (style.dash === "dot") return `${width} ${2 * width}`;
  if (style.dash === "dashDot") return `${4 * width} ${2 * width} ${width} ${2 * width}`;
  return undefined;
}

export function strokeAttributes(style: NormalizedStrokeStyle, theme: NormalizedPresentationTheme): string {
  const dash = dashArray(style);
  return `${paintAttributes(style.paint, theme, "stroke")} stroke-width="${style.width}"${dash === undefined ? "" : ` stroke-dasharray="${dash}"`} stroke-linecap="round" stroke-linejoin="round"`;
}

export function transformAttribute(box: Box, transform: NormalizedTransform): string {
  if (transform.rotation === 0 && !transform.flipH && !transform.flipV) return "";
  const cx = box.x + box.width / 2;
  const cy = box.y + box.height / 2;
  return ` transform="translate(${cx} ${cy}) rotate(${transform.rotation}) scale(${transform.flipH ? -1 : 1} ${transform.flipV ? -1 : 1}) translate(${-cx} ${-cy})"`;
}
