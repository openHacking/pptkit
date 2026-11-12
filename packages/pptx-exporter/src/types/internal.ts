import type { NormalizedAsset } from "@pptkit/core";
export interface ZipPart {
  name: string;
  data: Uint8Array;
}

export interface LoadedAsset {
  data: Uint8Array;
  mimeType: string;
  extension: string;
  animationFallback?: "first-frame";
  orientationNormalization?: "pixels";
  vectorRasterization?: "emf" | "wmf" | "svg";
  svgPair?: Uint8Array;
}

export interface PackagedMedia {
  name: string;
  loaded: LoadedAsset;
  svgName?: string;
}

export interface Relationship {
  id: string;
  type: string;
  target: string;
  targetMode?: "External";
}

export interface AssetLoadContext { signal?: AbortSignal; maxAssetBytes?: number; externalImageValidation?: "signature" }
export type AssetLoader = (asset: NormalizedAsset, context: AssetLoadContext) => Promise<LoadedAsset>;
