export interface ZipPart {
  name: string;
  data: Buffer;
}

export interface LoadedAsset {
  data: Buffer;
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
