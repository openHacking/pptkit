export interface ZipInspectionLimits {
  maxInputBytes: number;
  maxEntries: number;
  maxEntryBytes: number;
  maxExpandedBytes: number;
  maxCompressionRatio: number;
  /** Bounds compressed input accepted without producing any decoded bytes. */
  maxStalledInputBytes: number;
}
export interface ZipInspectionOptions {
  limits?: Partial<ZipInspectionLimits>;
  signal?: AbortSignal;
}
export interface ZipEntryInfo {
  readonly name: string;
  readonly directory: boolean;
  readonly method: "stored" | "deflate";
  readonly compressedBytes: number;
  readonly expandedBytes: number;
  readonly crc32: number;
}
export interface ZipArchiveInspection {
  readonly entries: readonly ZipEntryInfo[];
  readonly inputBytes: number;
  readonly expandedBytes: number;
}
export type PackageInspectionErrorCode = "archive-invalid" | "archive-unsupported" | "archive-path-invalid" | "archive-duplicate-path" | "archive-crc-mismatch" | "archive-limit-exceeded" | "operation-aborted" | "xml-invalid" | "xml-unsupported" | "xml-limit-exceeded" | "package-invalid" | "package-unsupported" | "package-limit-exceeded" | "presentation-invalid" | "presentation-unsupported" | "mce-invalid" | "mce-unsupported";
export interface PackageSourceLocation {
  /** One-based XML line and column; offset counts UTF-16 code units. */
  readonly line: number;
  readonly column: number;
  readonly offset: number;
}
export class PackageInspectionError extends Error {
  readonly code: PackageInspectionErrorCode;
  readonly entry: string | undefined;
  readonly location: PackageSourceLocation | undefined;
  constructor(code: PackageInspectionErrorCode, message: string, entry?: string, location?: PackageSourceLocation) {
    super(message); this.name = "PackageInspectionError"; this.code = code; this.entry = entry; this.location = location ? Object.freeze({ ...location }) : undefined;
  }
}
