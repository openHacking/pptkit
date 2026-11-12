export type PresentationAssetSource =
  | { type: "path" | "url"; value: string }
  | { type: "bytes"; value: string; encoding: "base64" };

/** Raw byte inputs are copied and encoded when registered; stored state uses Base64. */
export type PresentationAssetSourceInput = PresentationAssetSource
  | { type: "bytes"; value: Uint8Array; encoding?: "raw" };

export interface PresentationAssetInput {
  id?: string;
  kind: "image";
  source: PresentationAssetSourceInput;
  mimeType?: string;
  width?: number;
  height?: number;
  accessibility?: {
    description?: string;
    decorative?: boolean;
  };
  dedupeKey?: string;
}

export interface PresentationAsset extends Omit<PresentationAssetInput, "id" | "source"> {
  id: string;
  source: PresentationAssetSource;
}

export interface NormalizedAsset {
  id: string;
  kind: "image";
  source: PresentationAssetSource;
  mimeType?: string;
  width?: number;
  height?: number;
  accessibility: {
    description?: string;
    decorative: boolean;
  };
  dedupeKey?: string;
}


export interface AssetRegistrationOptions {
  /** Opt-in exact byte and metadata matching; explicit IDs always retain their identity. */
  deduplicate?: "content";
}
/** One item in an atomic batch registration. */
export interface AssetRegistrationRequest {
  readonly asset: PresentationAssetInput;
  readonly options?: AssetRegistrationOptions;
}

/** Explicit cross-document identity and external-source policies. */
export interface AssetCopyOptions {
  readonly deduplicate: "registry" | "content";
  readonly external: "reject" | "preserve";
}
export interface AssetCopyMapping {
  readonly sourceId: string;
  readonly destinationId: string;
  /** True when this destination identity existed before this item in the batch. */
  readonly reused: boolean;
}
export type PresentationAssetReplacement = Omit<PresentationAssetInput, "id">;
export interface AssetRetentionOptions {
  /** Additional Core asset IDs retained by a consumer outside the known image graph. */
  retainAssetIds?: readonly string[];
}
export interface AssetReference {
  readonly assetId: string;
  readonly elementId: string;
  readonly path: string;
  readonly slideId?: string;
  readonly layoutId?: string;
  readonly masterId?: string;
}
