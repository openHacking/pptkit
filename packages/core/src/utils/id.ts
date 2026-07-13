let presentationCounter = 0;

export function createDocumentId(): string {
  presentationCounter += 1;
  return `presentation-${presentationCounter}`;
}

export function createSlideId(index: number): string {
  return `slide-${index + 1}`;
}

export function createAssetId(index: number): string {
  return `asset-${index + 1}`;
}
