import type { ZipArchiveInspection, ZipInspectionLimits } from "./archive.js";
import type { XmlInspectionLimits, XmlPartInspection } from "./xml.js";
import type { MarkupCompatibilityReport } from "./compatibility.js";
export interface OpcInspectionLimits {
  maxXmlParts: number;
  maxXmlCharacters: number;
  maxXmlNodes: number;
  maxXmlTextCharacters: number;
  maxRelationships: number;
}
export interface OpcInspectionOptions {
  zipLimits?: Partial<ZipInspectionLimits>;
  xmlLimits?: Partial<XmlInspectionLimits>;
  limits?: Partial<OpcInspectionLimits>;
  signal?: AbortSignal;
}
export interface OpcPartInfo {
  /** Absolute logical part name; ZIP names are mapped separately. */
  readonly name: string;
  readonly entryName: string;
  readonly contentType: string;
  readonly relationshipPart: boolean;
  readonly expandedBytes: number;
  readonly xml?: XmlPartInspection;
  /** Effective relationship markup decisions when the source declares/uses MCE. */
  readonly markupCompatibility?: MarkupCompatibilityReport;
}
export interface OpcRelationshipInfo {
  readonly id: string;
  readonly type: string;
  /** '/' denotes the package; otherwise the actual logical source part name. */
  readonly source: string;
  readonly relationshipPart: string;
  readonly target: string;
  readonly targetMode: "Internal" | "External";
  /** Actual logical target name, present only for resolved internal relationships. */
  readonly targetPart?: string;
  /** Uninterpreted fragment spelling, without '#'. */
  readonly fragment?: string;
}
export interface OpcPackageInspection {
  readonly archive: ZipArchiveInspection;
  readonly parts: readonly OpcPartInfo[];
  readonly relationships: readonly OpcRelationshipInfo[];
  readonly contentTypes: {
    readonly defaults: readonly { readonly extension: string; readonly contentType: string }[];
    readonly overrides: readonly { readonly partName: string; readonly contentType: string }[];
  };
  /** Includes [Content_Types].xml, which is a stream rather than an OPC part. */
  readonly xmlTotals: { readonly parts: number; readonly characters: number; readonly nodes: number; readonly textCharacters: number };
}
