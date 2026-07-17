import type { DeckSessionV2 } from "./contracts.js";

export const SESSION_SCHEMA_VERSION = 2 as const;
const SUPPORTED_ASSET_MIME_TYPES = new Set(["image/png", "image/jpeg", "image/gif", "image/svg+xml"]);
const SHA256 = /^[a-f0-9]{64}$/i;
const COMPOSITIONS = new Set(["hero", "split", "ledger", "grid", "divided", "timeline", "image-split", "image-hero"]);
const DENSITIES = new Set(["airy", "balanced", "dense"]);
const VARIATIONS = new Set(["restrained", "balanced", "expressive"]);
const THEMES = new Set(["clean-business", "swiss-grid", "editorial-story"]);

export function parseDeckSession(value: string | unknown): DeckSessionV2 {
  const input: unknown = typeof value === "string" ? JSON.parse(value) : value;
  if (!input || typeof input !== "object") throw new Error("Deck session must be an object.");
  const candidate = input as Partial<DeckSessionV2>;
  if (candidate.schemaVersion !== SESSION_SCHEMA_VERSION) throw new Error(`Unsupported deck session schema: ${String(candidate.schemaVersion)}.`);
  if (!candidate.id || !candidate.deck || !Array.isArray(candidate.sources) || !Array.isArray(candidate.assets)) throw new Error("Deck session is missing required fields.");
  if (!candidate.deck.design || !candidate.deck.design.seed?.trim()) throw new Error("Deck session requires a non-empty design seed.");
  if (!THEMES.has(candidate.deck.design.theme?.id)) throw new Error(`Unsupported theme: ${String(candidate.deck.design.theme?.id)}.`);
  if (!VARIATIONS.has(candidate.deck.design.variation)) throw new Error(`Unsupported design variation: ${String(candidate.deck.design.variation)}.`);
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
    if (slide.composition && !COMPOSITIONS.has(slide.composition)) throw new Error(`Unsupported slide composition: ${slide.composition}.`);
    if (slide.density && !DENSITIES.has(slide.density)) throw new Error(`Unsupported slide density: ${slide.density}.`);
    if (slide.image && !assetIds.has(slide.image.assetId)) throw new Error(`Slide ${slide.id} references undeclared asset ${slide.image.assetId}.`);
  }
  return candidate as DeckSessionV2;
}

export function serializeDeckSession(session: DeckSessionV2) {
  parseDeckSession(session);
  return JSON.stringify(session, null, 2);
}
