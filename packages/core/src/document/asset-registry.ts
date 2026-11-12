import { captureAssetSource } from "./asset-source.js";
import type { AssetRegistrationOptions, PresentationAssetReplacement, PresentationAsset, PresentationAssetInput } from "../types/asset.js";
import { deepClone, deepFreeze } from "../utils/clone.js";
import { IdAllocator } from "../utils/id.js";

function sameAssetPayload(left: PresentationAsset, right: PresentationAsset): boolean {
  return left.id === right.id && left.kind === right.kind &&
    left.source.type === right.source.type && left.source.value === right.source.value &&
    left.mimeType === right.mimeType && Object.is(left.width, right.width) && Object.is(left.height, right.height) &&
    left.dedupeKey === right.dedupeKey &&
    (left.accessibility === undefined) === (right.accessibility === undefined) &&
    left.accessibility?.description === right.accessibility?.description && left.accessibility?.decorative === right.accessibility?.decorative;
}

export class AssetRegistry {
  private entries: PresentationAsset[] = [];
  private allocator = new IdAllocator("asset");
  private version = 0;

  get revision(): number { return this.version; }
  fork(): AssetRegistry {
    const fork = new AssetRegistry();
    fork.entries = [...this.entries];
    fork.allocator = this.allocator.fork();
    fork.version = this.version;
    return fork;
  }
  commitFrom(staged: AssetRegistry, expectedRevision: number): void {
    if (this.version !== expectedRevision) throw new Error("Asset registry changed during batch registration.");
    this.entries = staged.entries;
    this.allocator = staged.allocator;
    this.version = staged.version;
  }

  get assets(): readonly PresentationAsset[] {
    return Object.freeze([...this.entries]);
  }

  register(input: PresentationAssetInput, options: AssetRegistrationOptions = {}): PresentationAsset {
    if (!options || typeof options !== "object" || Array.isArray(options) || Object.keys(options).some(key => key !== "deduplicate")) throw new TypeError("Invalid asset registration options.");
    const { deduplicate } = options;
    if (deduplicate !== undefined && deduplicate !== "content") throw new TypeError("Unknown asset deduplication mode.");
    const fields = { ...input };
    const captured = deepClone({ ...fields, source: captureAssetSource(fields.source) });
    const existing = this.findByIdentity(captured);
    if (existing !== undefined) {
      const candidate = deepFreeze({ ...deepClone(captured), id: existing.id }) as PresentationAsset;
      if (!sameAssetPayload(existing, candidate)) {
        throw new Error(`Asset "${existing.id}" is already registered with different metadata.`);
      }
      return existing;
    }

    if (captured.id === undefined && deduplicate === "content" && captured.source.type === "bytes") {
      const matching = this.entries.find(asset => sameAssetPayload(asset, { ...captured, id: asset.id }));
      if (matching !== undefined) return matching;
    }

    const id = captured.id ?? this.allocator.next((candidate) => this.entries.some((asset) => asset.id === candidate));
    if (this.entries.some((asset) => asset.id === id)) {
      throw new Error(`Duplicate asset id "${id}" detected.`);
    }
    const asset = deepFreeze({ ...deepClone(captured), id }) as PresentationAsset;
    this.entries.push(asset);
    this.version++;
    return asset;
  }

  replace(assetId: string, replacement: PresentationAssetReplacement): PresentationAsset {
    const previous = this.get(assetId);
    if (!previous) throw new Error(`Unknown asset id "${assetId}".`);
    const fields = { ...replacement };
    if (Object.hasOwn(fields, "id")) throw new TypeError("Asset replacement cannot specify an ID; the existing identity is retained.");
    const candidate = deepFreeze(deepClone({ ...fields, source: captureAssetSource(fields.source), id: assetId })) as PresentationAsset;
    // Capture all caller values before mutation. Existing immutable views stay unchanged.
    const index = this.entries.indexOf(previous);
    if (index < 0) throw new Error(`Asset "${assetId}" changed while capturing its replacement.`);
    this.entries[index] = candidate;
    this.version++;
    return candidate;
  }

  remove(assetId: string): PresentationAsset {
    const index = this.entries.findIndex(asset => asset.id === assetId);
    if (index < 0) throw new Error(`Unknown asset id "${assetId}".`);
    const removed = this.entries.splice(index, 1)[0]!;
    this.version++;
    return removed;
  }

  removeExcept(retained: ReadonlySet<string>): readonly PresentationAsset[] {
    const removed: PresentationAsset[] = [];
    let write = 0;
    for (const asset of this.entries) {
      if (retained.has(asset.id)) this.entries[write++] = asset;
      else removed.push(asset);
    }
    this.entries.length = write;
    if (removed.length) this.version++;
    return Object.freeze(removed);
  }

  get(assetId: string): PresentationAsset | undefined {
    return this.entries.find((asset) => asset.id === assetId);
  }

  private findByIdentity(input: Omit<PresentationAssetInput, "source"> & { source: PresentationAsset["source"] }): PresentationAsset | undefined {
    if (input.id !== undefined) return this.get(input.id);
    if (input.dedupeKey !== undefined) {
      return this.entries.find((asset) => asset.dedupeKey === input.dedupeKey &&
        asset.source.type === input.source.type && asset.source.value === input.source.value);
    }
    return undefined;
  }
}
