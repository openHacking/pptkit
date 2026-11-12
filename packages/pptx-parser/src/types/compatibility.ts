import type { XmlExpandedName, XmlInspectionOptions, XmlPartInspection } from "./xml.js";

export interface MarkupCompatibilityOptions extends XmlInspectionOptions {
  /** Exact namespace URIs understood by the consuming application; not inferred from declarations. */
  understoodNamespaces: readonly string[];
  /** Whether unqualified elements/attributes are understood. Defaults to true. */
  understandNoNamespace?: boolean;
  /** Expanded names whose entire element and contents bypass MCE processing. */
  extensionElements?: readonly { namespaceUri: string; local: string }[];
}
export interface MarkupCompatibilityDecision {
  /** One-based element index in the original XML, including opaque/removed content. */
  readonly elementIndex: number;
  readonly name: XmlExpandedName;
  readonly action: "ignored-element" | "unwrapped-element" | "ignored-attribute" | "alternate-content" | "opaque-extension";
  readonly attribute?: XmlExpandedName;
  /** Selected Choice/Fallback source element index; absent when no branch matches. */
  readonly selectedElementIndex?: number;
}
export interface MarkupCompatibilityReport {
  /** May have zero/multiple roots after replacing a root AlternateContent. */
  readonly output: {
    readonly roots: readonly XmlExpandedName[];
    readonly elements: number;
    readonly attributes: number;
    /** Text/CDATA content only, in UTF-16 code units; excludes comments and attributes. */
    readonly textCharacters: number;
  };
  readonly decisions: readonly MarkupCompatibilityDecision[];
}
export interface MarkupCompatibilityInspection extends MarkupCompatibilityReport {
  readonly source: XmlPartInspection;
}
