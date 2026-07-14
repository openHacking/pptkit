export function escapeXml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

export function safeId(value: string): string {
  const normalized = value.replace(/[^A-Za-z0-9_.:-]/g, "-");
  let hash = 2166136261;
  for (let index = 0; index < value.length; index += 1) {
    hash ^= value.charCodeAt(index);
    hash = Math.imul(hash, 16777619);
  }
  const base = /^[A-Za-z_]/.test(normalized) ? normalized : `id-${normalized}`;
  return `${base}-${(hash >>> 0).toString(36)}`;
}

export function safeUrl(value: string): string | undefined {
  try {
    const url = new URL(value, "https://pptkit.invalid");
    if (["http:", "https:", "blob:"].includes(url.protocol)) return value;
    if (url.protocol === "data:" && /^data:image\//i.test(value)) return value;
  } catch {
    return undefined;
  }
  return undefined;
}

export function safeActionUrl(value: string): string | undefined {
  try {
    const url = new URL(value, "https://pptkit.invalid");
    if (["http:", "https:", "mailto:"].includes(url.protocol)) return value;
  } catch {
    return undefined;
  }
  return undefined;
}
