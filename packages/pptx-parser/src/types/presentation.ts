import type { OpcPackageInspection } from "./opc.js";

export type PresentationDialect = "strict" | "transitional";
export type PresentationFileKind = "presentation" | "template" | "slideshow";
/** Entry-point identification, not schema conformance or semantic import. */
export interface PresentationPackageInspection {
  readonly package: OpcPackageInspection;
  readonly mainPart: string;
  readonly mainRelationshipId: string;
  /** Dialect of the entry relationship and main root, which must agree. */
  readonly dialect: PresentationDialect;
  readonly kind: PresentationFileKind;
  /** Main content type permits macros; does not assert presence or absence of executable content. */
  readonly macroEnabled: boolean;
  readonly suggestedExtension: "pptx" | "pptm" | "potx" | "potm" | "ppsx" | "ppsm";
}
