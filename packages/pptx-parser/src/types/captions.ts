import type { PptxInspectionOptions } from './structure.js';
import type { PptxSourceLocation } from './source.js';

export interface PptxCaptionInspectionLimits {
  readonly maxWorkUnits: number;
  readonly maxTracks: number;
  readonly maxTextCharacters: number;
  readonly maxTrackBytes: number;
  readonly maxTotalTrackBytes: number;
}
export interface PptxCaptionInspectionOptions extends PptxInspectionOptions {
  readonly captionLimits?: Partial<PptxCaptionInspectionLimits>;
}
export interface PptxCaptionTrackSource {
  readonly slideId: number;
  readonly source: PptxSourceLocation;
  /** Picture presence only; the extension and track remain source-only. */
  readonly pictureView: 'effective-drawing' | 'raw-only';
  readonly objectId?: string;
  readonly trackId?: string;
  readonly language?: string;
  readonly label?: string;
  readonly relationshipId?: string;
  readonly status: 'vtt-header-only' | 'invalid-utf8' | 'invalid-vtt-header' | 'unsupported-content-type' | 'external' | 'unsupported-relationship';
  readonly targetPart?: string;
  readonly externalTarget?: string;
  readonly contentType?: string;
  readonly byteLength?: number;
}
export interface PptxCaptionInspection {
  readonly schemaVersion: 'pptkit-caption-source-v1';
  readonly dialect: 'strict' | 'transitional';
  readonly view: 'raw-only';
  /** Source-only records may include inactive markup compatibility branches. */
  readonly tracks: readonly PptxCaptionTrackSource[];
  readonly workUnits: number;
  readonly inspectedBytes: number;
}
