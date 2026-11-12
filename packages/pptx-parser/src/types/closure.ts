import type { OpcInspectionOptions } from "./opc.js";

export interface OpcClosureLimits {
  maxRoots: number;
  /** Counts all raw Relationship declarations, including unselected MCE branches. */
  maxSourceRelationships: number;
  /** Unique selected parts, including their relationship-set parts. */
  maxClosureParts: number;
  /** Unique selected expanded bytes plus the content-types control stream. */
  maxClosureExpandedBytes: number;
  /** All source relationship findings; exceeding this rejects, never truncates. */
  maxFindings: number;
  /** Raw XML traversal and graph indexing/traversal work. */
  maxWorkUnits: number;
}
export interface OpcClosureOptions extends OpcInspectionOptions {
  /** '/' means the package root; other roots must identify ordinary OPC parts. */
  roots: readonly string[];
  closureLimits?: Partial<OpcClosureLimits>;
}
export interface OpcSourceRelationship {
  readonly source: string;
  readonly relationshipPart: string;
  /** One-based preorder index in the raw relationship XML, including MCE wrappers. */
  readonly elementIndex: number;
  /** Missing/invalid fields in inactive source markup are reported by findings. */
  readonly id?: string;
  readonly type?: string;
  readonly target?: string;
  readonly targetMode?: "Internal" | "External";
  readonly targetPart?: string;
  readonly fragment?: string;
  /** Value match only: duplicate declarations in different branches may both match. */
  readonly matchesEffectiveRelationship: boolean;
}
export interface OpcClosureFinding {
  readonly code: "opaque-source-markup" | "invalid-source-relationship" | "missing-source-target" | "ambiguous-source-identity";
  readonly source: string;
  readonly relationshipPart: string;
  readonly elementIndex: number;
  readonly reason: string;
}
export interface OpcClosurePart {
  readonly name: string;
  readonly entryName: string;
  readonly contentType: string;
  readonly expandedBytes: number;
  readonly reason: "root" | "dependency" | "relationship-set";
  /** First deterministic discovery edge; roots/implicit relationship sets omit it. */
  readonly via?: {
    readonly source: string;
    readonly relationshipId?: string;
    readonly elementIndex: number;
  };
}
export interface OpcClosureInspection {
  readonly profile: "source-relationships-v1";
  readonly roots: readonly string[];
  readonly parts: readonly OpcClosurePart[];
  readonly relationships: readonly OpcSourceRelationship[];
  /** Source relationships outside this closure targeting selected ordinary parts. */
  readonly incomingRelationships: readonly OpcSourceRelationship[];
  /** Unreachable from these roots, never a garbage-collection authorization. */
  readonly unreachedParts: readonly string[];
  /** Includes findings outside the closure to explain incomplete incoming coverage. */
  readonly findings: readonly OpcClosureFinding[];
  readonly coverage: {
    /** Completeness only for recognized source relationship markup, not all OOXML references. */
    readonly outgoingComplete: boolean;
    readonly incomingComplete: boolean;
  };
  readonly controlStream: {
    readonly name: "/[Content_Types].xml";
    readonly expandedBytes: number;
  };
  readonly expandedBytes: number;
  readonly sourceRelationshipCount: number;
  readonly effectiveRelationshipCount: number;
  readonly workUnits: number;
}
