import type { PresentationAsset, PresentationAssetInput } from "../types/asset.js";
import { cloneAsset } from "../utils/clone.js";
import { createAssetId } from "../utils/id.js";
import { assertOptionalDimension } from "../validation/geometry.js";

function sameAssetPayload(left: PresentationAsset, right: PresentationAsset): boolean {
  return left.kind === right.kind && left.source.type === right.source.type &&
    left.source.value === right.source.value && left.mimeType === right.mimeType &&
    left.width === right.width && left.height === right.height && left.altText === right.altText &&
    left.dedupeKey === right.dedupeKey;
}

export class AssetRegistry {
  readonly assets: PresentationAsset[] = [];

  register(input: PresentationAssetInput): PresentationAsset {
    const existing = this.findByIdentity(input);
    if (existing !== undefined) {
      const normalizedInput = cloneAsset({ ...input, id: existing.id });
      if (!sameAssetPayload(existing, normalizedInput)) {
        throw new Error(`Asset "${existing.id}" is already registered with different metadata.`);
      }
      return existing;
    }

    const asset = cloneAsset({ ...input, id: input.id ?? createAssetId(this.assets.length) });
    if (this.assets.some((existingAsset) => existingAsset.id === asset.id)) {
      throw new Error(`Duplicate asset id "${asset.id}" detected.`);
    }
    assertOptionalDimension(asset.width, `Asset "${asset.id}" width`);
    assertOptionalDimension(asset.height, `Asset "${asset.id}" height`);
    this.assets.push(asset);
    return asset;
  }

  get(assetId: string): PresentationAsset | undefined {
    return this.assets.find((asset) => asset.id === assetId);
  }

  private findByIdentity(input: PresentationAssetInput): PresentationAsset | undefined {
    if (input.id !== undefined) return this.get(input.id);
    if (input.dedupeKey !== undefined) {
      return this.assets.find((asset) => asset.dedupeKey === input.dedupeKey &&
        asset.source.type === input.source.type && asset.source.value === input.source.value);
    }
    return undefined;
  }
}
