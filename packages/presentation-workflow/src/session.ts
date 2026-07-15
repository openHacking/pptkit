import type { DeckSessionV1 } from "./contracts.js";

export const SESSION_SCHEMA_VERSION = 1 as const;
export const MAX_INLINE_ASSET_BYTES = 5 * 1024 * 1024;
export const MAX_INLINE_ASSETS_BYTES = 20 * 1024 * 1024;

function approximateDataUrlBytes(value: string) {
  const comma = value.indexOf(",");
  if (comma < 0) return value.length;
  return Math.ceil((value.length - comma - 1) * 0.75);
}

export function parseDeckSession(value: string | unknown): DeckSessionV1 {
  const input: unknown = typeof value === "string" ? JSON.parse(value) : value;
  if (!input || typeof input !== "object") throw new Error("Deck session must be an object.");
  const candidate = input as Partial<DeckSessionV1>;
  if (candidate.schemaVersion !== SESSION_SCHEMA_VERSION) throw new Error(`Unsupported deck session schema: ${String(candidate.schemaVersion)}.`);
  if (!candidate.id || !candidate.deck || !Array.isArray(candidate.sources) || !Array.isArray(candidate.assets)) throw new Error("Deck session is missing required fields.");
  let total = 0;
  for (const asset of candidate.assets) {
    if (!asset?.id || !asset.name || !asset.mimeType) throw new Error("Every session asset requires id, name, and mimeType.");
    if (!asset.dataUrl) continue;
    const size = approximateDataUrlBytes(asset.dataUrl);
    if (size > MAX_INLINE_ASSET_BYTES) throw new Error(`Inline asset ${asset.name} exceeds the 5 MB limit; choose the file manually or use the Node fallback.`);
    total += size;
  }
  if (total > MAX_INLINE_ASSETS_BYTES) throw new Error("Inline session assets exceed the 20 MB limit; choose files manually or use the Node fallback.");
  return candidate as DeckSessionV1;
}

export function serializeDeckSession(session: DeckSessionV1) {
  parseDeckSession(session);
  return JSON.stringify(session, null, 2);
}
