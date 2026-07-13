export interface PresentationAssetSource {
  type: "path" | "url";
  value: string;
}

export interface PresentationAssetInput {
  id?: string;
  kind: "image";
  source: PresentationAssetSource;
  mimeType?: string;
  width?: number;
  height?: number;
  altText?: string;
  dedupeKey?: string;
}

export interface PresentationAsset {
  id: string;
  kind: "image";
  source: PresentationAssetSource;
  mimeType?: string;
  width?: number;
  height?: number;
  altText?: string;
  dedupeKey?: string;
}

export interface NormalizedAsset extends PresentationAsset {}
