export interface XmlInspectionLimits {
  maxInputBytes: number;
  /** Decoded source length in UTF-16 code units, excluding the encoding BOM. */
  maxCharacters: number;
  /** Elements, coalesced text runs, CDATA sections, comments and processing instructions. */
  maxNodes: number;
  /** Root depth is one. */
  maxDepth: number;
  /** Includes namespace declaration attributes. */
  maxAttributesPerElement: number;
  /** Total decoded text, CDATA, attribute values, comments and PI bodies. */
  maxTextCharacters: number;
  /** Maximum source span between parser events; bounds unfinished token buffering. */
  maxPendingCharacters: number;
}
export interface XmlInspectionOptions {
  limits?: Partial<XmlInspectionLimits>;
  signal?: AbortSignal;
  /** Diagnostic label only. Never interpreted as a URL or filesystem path. */
  partName?: string;
}
export interface XmlExpandedName {
  readonly name: string;
  readonly prefix: string;
  readonly local: string;
  readonly namespaceUri: string;
}
export interface XmlPartInspection {
  readonly root: XmlExpandedName;
  readonly encoding: "utf-8" | "utf-16le" | "utf-16be";
  readonly inputBytes: number;
  readonly characters: number;
  readonly nodes: number;
  readonly elements: number;
  readonly attributes: number;
  readonly maxDepth: number;
  readonly textCharacters: number;
  /** Nonempty namespace URIs encountered in declarations, element or attribute names. */
  readonly namespaceUris: readonly string[];
}
