import type { PresentationAsset, PresentationAssetInput } from "../types/asset.js";
import { deepClone, deepFreeze } from "../utils/clone.js";
import { IdAllocator } from "../utils/id.js";

function sameAssetPayload(left: PresentationAsset, right: PresentationAsset): boolean {
  return JSON.stringify(left) === JSON.stringify(right);
}

export class AssetRegistry {
  private readonly entries: PresentationAsset[] = [];
  private readonly allocator = new IdAllocator("asset");

  get assets(): readonly PresentationAsset[] {
    return Object.freeze([...this.entries]);
  }

  register(input: PresentationAssetInput): PresentationAsset {
    const existing = this.findByIdentity(input);
    if (existing !== undefined) {
      const candidate = deepFreeze({ ...deepClone(input), id: existing.id }) as PresentationAsset;
      if (!sameAssetPayload(existing, candidate)) {
        throw new Error(`Asset "${existing.id}" is already registered with different metadata.`);
      }
      return existing;
    }

    const id = input.id ?? this.allocator.next((candidate) => this.entries.some((asset) => asset.id === candidate));
    if (this.entries.some((asset) => asset.id === id)) {
      throw new Error(`Duplicate asset id "${id}" detected.`);
    }
    const asset = deepFreeze({ ...deepClone(input), id }) as PresentationAsset;
    this.entries.push(asset);
    return asset;
  }

  get(assetId: string): PresentationAsset | undefined {
    return this.entries.find((asset) => asset.id === assetId);
  }

  private findByIdentity(input: PresentationAssetInput): PresentationAsset | undefined {
    if (input.id !== undefined) return this.get(input.id);
    if (input.dedupeKey !== undefined) {
      return this.entries.find((asset) => asset.dedupeKey === input.dedupeKey &&
        asset.source.type === input.source.type && asset.source.value === input.source.value);
    }
    return undefined;
  }
}
