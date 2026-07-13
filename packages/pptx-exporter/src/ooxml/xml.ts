import { EMU_PER_PT } from "../constants/ooxml.js";

export function emu(value: number): number {
  return Math.round(value * EMU_PER_PT);
}

export function colorValue(value: string): string {
  return value.replace(/^#/, "").padStart(6, "0").slice(-6).toUpperCase();
}

export function groupTransformXml(): string {
  return `<a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm>`;
}

export function paragraphAlignment(value: string | undefined): "l" | "ctr" | "r" | "just" {
  if (value === "center") return "ctr";
  if (value === "right") return "r";
  if (value === "justify") return "just";
  return "l";
}

export function escapeXml(value: string): string {
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");
}
