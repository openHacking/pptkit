import type { DeckSessionV1 } from "./contracts.js";

export const SESSION_SCHEMA_VERSION = 1 as const;
const SUPPORTED_ASSET_MIME_TYPES = new Set(["image/png", "image/jpeg", "image/gif", "image/svg+xml"]);
const SHA256 = /^[a-f0-9]{64}$/i;

export function parseDeckSession(value: string | unknown): DeckSessionV1 {
  const input: unknown = typeof value === "string" ? JSON.parse(value) : value;
  if (!input || typeof input !== "object") throw new Error("Deck session must be an object.");
  const candidate = input as Partial<DeckSessionV1>;
  if (candidate.schemaVersion !== SESSION_SCHEMA_VERSION) throw new Error(`Unsupported deck session schema: ${String(candidate.schemaVersion)}.`);
  if (!candidate.id || !candidate.deck || !Array.isArray(candidate.sources) || !Array.isArray(candidate.assets)) throw new Error("Deck session is missing required fields.");
  const assetIds = new Set<string>();
  for (const asset of candidate.assets) {
    if (!asset?.id || !asset.name || !asset.mimeType) throw new Error("Every session asset requires id, name, and mimeType.");
    if ("dataUrl" in asset) throw new Error(`Session asset ${asset.name} must not contain inline dataUrl content.`);
    if (!Number.isSafeInteger(asset.byteLength) || asset.byteLength <= 0) throw new Error(`Session asset ${asset.name} requires a positive byteLength.`);
    if (!SHA256.test(asset.sha256)) throw new Error(`Session asset ${asset.name} requires a valid SHA-256 digest.`);
    if (!SUPPORTED_ASSET_MIME_TYPES.has(asset.mimeType)) throw new Error(`Unsupported session asset MIME type: ${asset.mimeType}.`);
    if (assetIds.has(asset.id)) throw new Error(`Duplicate session asset id: ${asset.id}.`);
    assetIds.add(asset.id);
  }
  if (!Array.isArray(candidate.deck.slides)) throw new Error("Deck session slides must be an array.");
  for (const slide of candidate.deck.slides) {
    if (slide.image && !assetIds.has(slide.image.assetId)) throw new Error(`Slide ${slide.id} references undeclared asset ${slide.image.assetId}.`);
  }
  return candidate as DeckSessionV1;
}

export function serializeDeckSession(session: DeckSessionV1) {
  parseDeckSession(session);
  return JSON.stringify(session, null, 2);
}
