import type { SvgEmbeddedFont } from "../types/public.js";
import { cssFontFamily } from "./escape.js";

const alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
const maxFontBytes = 8 * 1024 * 1024;
const signatures: Record<SvgEmbeddedFont["format"], string> = { woff2: "wOF2", woff: "wOFF", truetype: "\x00\x01\x00\x00", opentype: "OTTO" };
const media: Record<SvgEmbeddedFont["format"], string> = { woff2: "font/woff2", woff: "font/woff", truetype: "font/ttf", opentype: "font/otf" };
export function embeddedFontMediaType(format: SvgEmbeddedFont["format"]): string { return media[format]; }

/** Snapshot and validate caller-provided programs without decoding the complete font. */
export function captureEmbeddedFonts(input: readonly SvgEmbeddedFont[] | undefined): { faces: SvgEmbeddedFont[]; style: string } {
  if (input === undefined) return { faces: [], style: "" };
  if (!Array.isArray(input) || input.length > 32) throw new TypeError("embeddedFonts must contain at most 32 font faces.");
  const faces: SvgEmbeddedFont[] = [], keys = new Set<string>();
  let total = 0;
  for (const entry of input as readonly SvgEmbeddedFont[]) {
    if (!entry || typeof entry !== "object") throw new TypeError("Each embedded font must be a font face object.");
    const { family, weight, italic, format, base64 } = entry;
    if (typeof family !== "string" || !family.trim() || /[\u0000-\u001f\u007f\ud800-\udfff\ufffe\uffff]/u.test(family)) throw new TypeError("Embedded font family must be a nonempty XML-safe name without controls.");
    if ((weight !== 400 && weight !== 700) || typeof italic !== "boolean" || !Object.hasOwn(signatures, format)) throw new TypeError("Embedded font style or format is invalid.");
    if (typeof base64 !== "string" || base64.length === 0 || base64.length > Math.ceil(maxFontBytes / 3) * 4 || base64.length % 4 !== 0 || !/^[A-Za-z0-9+/]*={0,2}$/.test(base64) || (base64.endsWith("==") && (alphabet.indexOf(base64[base64.length - 3]!) & 15) !== 0) || (base64.endsWith("=") && !base64.endsWith("==") && (alphabet.indexOf(base64[base64.length - 2]!) & 3) !== 0)) throw new TypeError("Embedded font data must be bounded canonical Base64.");
    const size = base64.length / 4 * 3 - (base64.endsWith("==") ? 2 : base64.endsWith("=") ? 1 : 0);
    total += size;
    if (total > maxFontBytes) throw new RangeError("Embedded font bytes exceed the 8 MiB render limit.");
    if (atob(base64.slice(0, 8)).slice(0, 4) !== signatures[format]) throw new TypeError("Embedded font signature does not match its format.");
    const key = JSON.stringify([family.toLowerCase(), weight, italic]);
    if (keys.has(key)) throw new TypeError("Embedded font family and style must be unique.");
    keys.add(key);
    faces.push({ family, weight, italic, format, base64 });
  }
  const declarations = faces.map(face => `@font-face{font-family:${cssFontFamily(face.family)};src:url(data:${media[face.format]};base64,${face.base64}) format('${face.format}');font-weight:${face.weight};font-style:${face.italic ? "italic" : "normal"}}`).join("");
  return { faces, style: declarations ? `<style>${declarations}</style>` : "" };
}
