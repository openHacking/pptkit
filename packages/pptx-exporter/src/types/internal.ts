export interface ZipPart {
  name: string;
  data: Uint8Array;
}

export interface LoadedAsset {
  data: Uint8Array;
  mimeType: string;
  extension: string;
}

export interface PackagedMedia {
  name: string;
  loaded: LoadedAsset;
}

export interface Relationship {
  id: string;
  type: string;
  target: string;
}

export type AssetLoader = (
  source: { type: "path" | "url"; value: string },
  declaredMime?: string,
) => Promise<LoadedAsset>;
