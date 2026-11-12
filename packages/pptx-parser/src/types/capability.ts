import type { ImportCapabilityTier, ImportFinding, ImportOperation, ImportReport } from "./import.js";
import type { PptxSourceLocation } from "./source.js";
export interface ImportCapabilityLimits {
  readonly maxFindings?: number;
  readonly maxTextCodeUnits?: number;
  readonly maxWorkUnits?: number;
  /** Complete matching subjects retained. Zero returns counts only. */
  readonly maxSubjects?: number;
}
export interface ImportCapabilityOptions {
  readonly limits?: ImportCapabilityLimits;
  /** Exact string matches, combined with AND; no regex or source-path evaluation. */
  readonly filter?: { readonly feature?: string; readonly part?: string; readonly slideId?: string; readonly elementId?: string };
}
export interface ImportCapabilityReferences {
  readonly rule: string;
  readonly match: "exact" | "prefix";
  readonly tasks: readonly string[];
  /** Repository-relative references, not proof these tests ran for this input. */
  readonly artifacts: readonly { readonly kind: "contract" | "test" | "native-observation"; readonly path: string }[];
}
export interface ImportCapabilityAssessment {
  readonly operation: ImportOperation;
  readonly status: "reported" | "unreported" | "conflicting";
  /** Present only when every observation for this exact subject/operation agrees. */
  readonly tier?: ImportCapabilityTier;
  readonly observations: readonly Pick<ImportFinding, "id" | "tier" | "reason">[];
}
export interface ImportCapabilitySubject {
  /** Ordinal within the complete sorted matching set; not a durable editing identity. */
  readonly id: string;
  readonly feature: string;
  readonly source: PptxSourceLocation;
  readonly slideId?: string;
  readonly elementId?: string;
  readonly operations: readonly ImportCapabilityAssessment[];
  readonly references?: ImportCapabilityReferences;
}
export interface ImportCapabilityReport {
  readonly schemaVersion: "pptkit-import-capabilities-v1";
  readonly catalogVersion: "pptkit-capability-references-v1";
  readonly coverage: "reported-source-findings";
  readonly projection: ImportReport["projection"];
  readonly nativeVerification: "not-performed";
  readonly candidateSaveValidation: "not-performed";
  readonly filter: NonNullable<ImportCapabilityOptions["filter"]>;
  readonly totalFindings: number;
  readonly matchedFindings: number;
  readonly totalSubjects: number;
  readonly returnedSubjects: number;
  readonly returnedFindings: number;
  readonly truncated: boolean;
  /** Counts over every matching subject, including omitted subjects. */
  readonly totals: readonly { readonly operation: ImportOperation; readonly reported: number; readonly unreported: number; readonly conflicting: number }[];
  readonly subjects: readonly ImportCapabilitySubject[];
  readonly workUnits: number;
}
