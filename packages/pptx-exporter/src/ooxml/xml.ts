import type { Box, ColorValue, NormalizedPaint, NormalizedStrokeStyle, NormalizedTransform } from "@pptkit/core";
import { EMU_PER_PT } from "../constants/ooxml.js";

export function emu(value: number): number {
  return Math.round(value * EMU_PER_PT);
}

export function colorValue(value: string): string {
  return value.replace(/^#/, "").toUpperCase();
}

const SCHEME_COLORS = {
  background1: "bg1",
  background2: "bg2",
  text1: "tx1",
  text2: "tx2",
  accent1: "accent1",
  accent2: "accent2",
  accent3: "accent3",
  accent4: "accent4",
  accent5: "accent5",
  accent6: "accent6",
  hyperlink: "hlink",
  followedHyperlink: "folHlink",
} as const;

export function colorXml(value: ColorValue, opacity = 1): string {
  const alpha = opacity < 1 ? `<a:alpha val="${Math.round(opacity * 100000)}"/>` : "";
  if (typeof value === "string") return `<a:srgbClr val="${colorValue(value)}">${alpha}</a:srgbClr>`;
  return `<a:schemeClr val="${SCHEME_COLORS[value.theme]}">${alpha}</a:schemeClr>`;
}

export function paintXml(paint: NormalizedPaint, opacity = 1): string {
  if (paint.type === "none") return "<a:noFill/>";
  return `<a:solidFill>${colorXml(paint.color, paint.opacity * opacity)}</a:solidFill>`;
}

function arrowXml(position: "headEnd" | "tailEnd", value: NormalizedStrokeStyle["beginArrow"]): string {
  return value === "none" ? "" : `<a:${position} type="${value}"/>`;
}

export function strokeXml(stroke: NormalizedStrokeStyle, opacity = 1, tag = "a:ln"): string {
  const dash = stroke.dash === "dot" ? "sysDot" : stroke.dash;
  return `<${tag} w="${emu(stroke.width)}">${paintXml(stroke.paint, opacity)}<a:prstDash val="${dash}"/>${arrowXml("headEnd", stroke.beginArrow)}${arrowXml("tailEnd", stroke.endArrow)}</${tag}>`;
}

function transformAttributes(transform: NormalizedTransform): string {
  return `${transform.rotation === 0 ? "" : ` rot="${Math.round(transform.rotation * 60000)}"`}${transform.flipH ? ' flipH="1"' : ""}${transform.flipV ? ' flipV="1"' : ""}`;
}

export function transformXml(box: Box, transform: NormalizedTransform): string {
  return `<a:xfrm${transformAttributes(transform)}><a:off x="${emu(box.x)}" y="${emu(box.y)}"/><a:ext cx="${emu(box.width)}" cy="${emu(box.height)}"/></a:xfrm>`;
}

export function groupTransformXml(box?: Box, transform?: NormalizedTransform, childWidth?: number, childHeight?: number): string {
  if (box === undefined || transform === undefined) return `<a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm>`;
  return `<a:xfrm${transformAttributes(transform)}><a:off x="${emu(box.x)}" y="${emu(box.y)}"/><a:ext cx="${emu(box.width)}" cy="${emu(box.height)}"/><a:chOff x="0" y="0"/><a:chExt cx="${emu(childWidth ?? box.width)}" cy="${emu(childHeight ?? box.height)}"/></a:xfrm>`;
}

export function paragraphAlignment(value: string): "l" | "ctr" | "r" | "just" {
  if (value === "center") return "ctr";
  if (value === "right") return "r";
  if (value === "justify") return "just";
  return "l";
}

export function escapeXml(value: string): string {
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");
}
