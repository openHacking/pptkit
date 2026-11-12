import type { OpcInspectionOptions } from "./opc.js";

export type OpcComparisonProfile = "opc-payload-v1" | "opc-xml-v1";
export interface OpcComparisonLimits {
  /** Total expanded bytes of both inputs, including control XML. */
  maxTotalExpandedBytes: number;
  /** Byte/string code units and XML tokens actually examined by comparison. */
  maxWorkUnits: number;
  /** Maximum retained findings; counting/validation continue after truncation. */
  maxDifferences: number;
}
export interface OpcComparisonOptions extends OpcInspectionOptions {
  /** Defaults to exact part payloads; XML normalization requires explicit opt-in. */
  profile?: OpcComparisonProfile;
  comparisonLimits?: Partial<OpcComparisonLimits>;
}
export type OpcDifferenceCode =
  | "part-added" | "part-removed" | "content-type-changed"
  | "payload-changed" | "xml-changed"
  | "relationship-added" | "relationship-removed" | "relationship-changed";
export interface OpcPackageDifference {
  readonly code: OpcDifferenceCode;
  /** Logical part name, or /[Content_Types].xml for the control stream. */
  readonly part: string;
  readonly relationshipId?: string;
  /** First unequal expanded byte, or the length of the shorter payload. */
  readonly byteOffset?: number;
  /** Parser-owned XML position, never an XPath expression or raw source handle. */
  readonly path?: string;
}
export interface OpcComparisonSummary {
  readonly parts: number;
  readonly relationships: number;
  readonly expandedBytes: number;
}
export interface OpcPackageComparison {
  readonly profile: OpcComparisonProfile;
  /** Equality under this profile, not OOXML validity, editability or visual parity. */
  readonly equal: boolean;
  readonly before: OpcComparisonSummary;
  readonly after: OpcComparisonSummary;
  /** Counts finding categories/parts/relationships, not every changed XML property. */
  readonly differenceCount: number;
  readonly differences: readonly OpcPackageDifference[];
  readonly truncated: boolean;
  readonly workUnits: number;
}
