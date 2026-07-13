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
  accessibility?: {
    description?: string;
    decorative?: boolean;
  };
  dedupeKey?: string;
}

export interface PresentationAsset extends Omit<PresentationAssetInput, "id"> {
  id: string;
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
