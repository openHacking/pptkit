/** Bounded prefix scan, not a PNG/APNG validator or image decoder. */
export function probePngAnimation(base64: string): "declared" | "not-declared" | "incomplete" {
  const maxBytes = 8 * 1024 * 1024, length = base64.length / 4 * 3 - (base64.endsWith("==") ? 2 : base64.endsWith("=") ? 1 : 0);
  if (!Number.isSafeInteger(length) || length > maxBytes) return "incomplete";
  let data: string; try { data = atob(base64); } catch { return "incomplete"; }
  if (!data.startsWith("\x89PNG\r\n\x1a\n")) return "incomplete";
  const u32 = (i: number) => data.charCodeAt(i) * 0x1000000 + data.charCodeAt(i+1) * 65536 + data.charCodeAt(i+2) * 256 + data.charCodeAt(i+3);
  let at = 8, chunks = 0;
  while (at < data.length && ++chunks <= 100000) {
    if (data.length - at < 12) return "incomplete";
    const size = u32(at), type = data.slice(at+4,at+8);
    if (size > 0x7fffffff || size > data.length - at - 12) return "incomplete";
    if (chunks === 1 && (type !== "IHDR" || size !== 13)) return "incomplete";
    if (type === "acTL") return size === 8 && u32(at+8) > 0 ? "declared" : "incomplete";
    if (type === "IDAT") return "not-declared";
    if (type === "IEND") return "incomplete";
    at += size + 12;
  }
  return "incomplete";
}
