/** Association inventory only; no node model, layout interpretation or editing permission. */
export type PptxDiagramComponentRole = "data" | "layout" | "colors" | "style";
export interface PptxDiagramComponentInfo {
  readonly role: PptxDiagramComponentRole;
  readonly status: "located" | "missing-reference" | "missing-relationship" | "wrong-relationship-type" | "external" | "fragment-target" | "wrong-content-type" | "wrong-root";
  readonly relationshipId?: string;
  /** Raw relationship target; never fetched. */
  readonly target?: string;
  readonly part?: string;
  readonly contentType?: string;
  /** Uninterpreted layoutDef/@uniqueId, not a supported layout/family declaration. */
  readonly layoutIdentifier?: string;
}
export interface PptxDiagramInfo {
  /** Raw source graphicFrame preorder identity, one-based within its slide XML. */
  readonly elementIndex: number;
  /** Raw source group preorder identities, outermost first. */
  readonly groupElementIndices: readonly number[];
  /** Lexical cNvPr identity/name, never a Core ID or inferred title. */
  readonly drawingId?: string;
  readonly name?: string;
  /** Unique is scoped to the effective slide shape tree, not the entire package. */
  readonly identityStatus: "unique" | "missing" | "invalid" | "ambiguous";
  readonly components: readonly PptxDiagramComponentInfo[];
  readonly cache?: PptxDiagramCacheInfo;
}

export interface PptxDiagramCacheReference {
  /** Raw data-part XML identity of the recognized extension block. */
  readonly elementIndex: number;
  readonly relationshipId?: string;
  /** Uninterpreted minimum application version URI. */
  readonly minimumVersion?: string;
}
export interface PptxDiagramCacheInfo {
  readonly status: PptxDiagramComponentInfo["status"] | "unresolved-data" | "ambiguous-reference" | "source-only";
  /** Presence/association never proves that data and cached drawing agree. */
  readonly freshness: "unknown";
  readonly dataPart?: string;
  readonly references: readonly PptxDiagramCacheReference[];
  readonly relationshipId?: string;
  readonly target?: string;
  readonly part?: string;
  readonly contentType?: string;
}
export interface PptxDiagramCacheCandidate {
  readonly slidePart: string;
  readonly relationshipId: string;
  readonly target: string;
  readonly part?: string;
  readonly contentType?: string;
  readonly status: "located" | "external" | "fragment-target" | "wrong-content-type" | "wrong-root";
  /** Located explicit associations only; empty does not mean safe to remove. */
  readonly referencedBy: readonly number[];
}
export interface PptxDiagramImageRelationship {
  readonly relationshipId: string;
  readonly status: "located" | "external" | "fragment-target" | "wrong-content-type" | "missing-relationship" | "wrong-relationship-type";
  readonly target?: string;
  readonly part?: string;
  readonly contentType?: string;
  /** Raw source blips, including inactive/opaque branches; not effective preview usage. */
  readonly references: readonly { readonly elementIndex: number; readonly attribute: "embed" | "link" }[];
}
export interface PptxDiagramResourceInfo {
  readonly part: string;
  readonly contentType: string;
  readonly roles: readonly (PptxDiagramComponentRole | "cache" | "image")[];
  /** Direct recognized associations. Image ownership is traced through imageRelationships. */
  readonly owners: readonly { readonly slidePart: string; readonly elementIndex: number; readonly role: PptxDiagramComponentRole | "cache" }[];
  /** All incoming edges in the effective OPC inventory, not a complete raw-source closure. */
  readonly incomingRelationships: readonly { readonly source: string; readonly relationshipId: string; readonly type: string }[];
  readonly imageRelationships: readonly PptxDiagramImageRelationship[];
  /** Raw cache shapes, including inactive branches. Repeated model IDs are legal here. */
  readonly modelReferences?: readonly { readonly elementIndex: number; readonly modelId?: string }[];
}
export interface PptxDiagramResourceInventory {
  readonly coverage: "direct-source-associations";
  readonly cacheCandidates: readonly PptxDiagramCacheCandidate[];
  readonly parts: readonly PptxDiagramResourceInfo[];
  readonly dataModels?: readonly PptxDiagramDataModelInfo[];
  /** One record per explicitly used data/cache pair; never a cross-product of resources. */
  readonly cacheMappings?: readonly PptxDiagramCacheMapping[];
  /** Opt-in direct paragraph observations; all cache freshness remains unknown. */
  readonly textComparisons?: readonly PptxDiagramTextComparison[];
}

export interface PptxDiagramModelReference {
  readonly value?: string;
  readonly status: "absent" | "unrecognized" | "missing" | "ambiguous" | "wrong-kind" | "resolved" | "uninterpreted";
  readonly targetElementIndex?: number;
}
export interface PptxDiagramTextParagraph {
  readonly elementIndex: number;
  /** Raw list-level spelling. This is source text, not resolved typography. */
  readonly level?: string;
  readonly runs: readonly { readonly elementIndex: number; readonly kind: "text" | "break" | "field"; readonly text: string; readonly fieldId?: string; readonly fieldType?: string }[];
}
export interface PptxDiagramPointInfo {
  readonly elementIndex: number;
  readonly modelId?: string;
  readonly identityStatus: "unique" | "duplicate" | "missing" | "unrecognized";
  readonly type?: string;
  readonly kind: "node" | "asst" | "doc" | "pres" | "parTrans" | "sibTrans" | "unrecognized";
  readonly connection?: PptxDiagramModelReference;
  readonly presentationAssociation?: PptxDiagramModelReference;
  /** Direct prSet attributes only; values are not evaluated as layout or theme parameters. */
  readonly properties?: readonly { readonly namespaceUri: string; readonly local: string; readonly value: string }[];
  readonly text?: readonly PptxDiagramTextParagraph[];
}
export interface PptxDiagramConnectionInfo {
  readonly elementIndex: number;
  readonly modelId?: string;
  readonly identityStatus: PptxDiagramPointInfo["identityStatus"];
  readonly type?: string;
  readonly kind: "parOf" | "presOf" | "presParOf" | "unknownRelationship" | "unrecognized";
  readonly source: PptxDiagramModelReference;
  readonly destination: PptxDiagramModelReference;
  readonly parentTransition?: PptxDiagramModelReference;
  readonly siblingTransition?: PptxDiagramModelReference;
  readonly sourceOrder?: string;
  readonly destinationOrder?: string;
  readonly sourceOrderValue?: number;
  readonly destinationOrderValue?: number;
  /** Layout/presentation URI spelling, not a point reference. */
  readonly presentationId?: string;
}
export interface PptxDiagramModelFinding {
  readonly code: "source-only" | "missing-field" | "unrecognized-field" | "duplicate-identity" | "unresolved-reference" | "uninterpreted-structure";
  readonly elementIndex: number;
  readonly attribute?: string;
  readonly reason: string;
}
export interface PptxDiagramDataModelInfo {
  readonly part: string;
  readonly coverage: "identities-references-direct-text";
  readonly status: "inspected" | "partial" | "source-only";
  readonly points: readonly PptxDiagramPointInfo[];
  readonly connections: readonly PptxDiagramConnectionInfo[];
  readonly findings: readonly PptxDiagramModelFinding[];
  /** Recognized explicit evidence only; no inferred text ownership or layout evaluation. */
  readonly presentationMappings?: readonly PptxDiagramPresentationMapping[];
}

export interface PptxDiagramPresentationMapping {
  /** Presentation point's raw data-part identity. */
  readonly elementIndex: number;
  readonly sources: readonly { readonly kind: "presAssocID" | "presOf"; readonly elementIndex: number; readonly reference: PptxDiagramModelReference }[];
  /** Describes recognized evidence; multiple targets do not by themselves imply an invalid diagram. */
  readonly agreement: "none" | "same-target" | "different-targets" | "unresolved";
  readonly targetElementIndex?: number;
}
export interface PptxDiagramCacheMapping {
  readonly dataPart: string;
  readonly cachePart: string;
  readonly coverage: "explicit-model-references";
  readonly status: "inspected" | "partial" | "source-only";
  readonly shapes: readonly { readonly elementIndex: number; readonly model: PptxDiagramModelReference }[];
  /** Findings are scoped to cachePart; referenced targets are scoped to dataPart. */
  readonly findings: readonly PptxDiagramModelFinding[];
}

export interface PptxDiagramTextComparison {
  readonly dataPart: string;
  readonly cachePart: string;
  readonly coverage: "direct-paragraph-vectors";
  readonly freshness: "unknown";
  readonly status: "observed" | "source-only";
  readonly reason?: string;
  readonly shapes: readonly {
    /** Raw index within cachePart. Source and presentation indices belong to dataPart. */
    readonly cacheElementIndex: number;
    readonly presentationElementIndex?: number;
    readonly sourcePointElementIndex?: number;
    readonly status: "equal" | "different" | "unavailable";
    readonly reason?: string;
    readonly sourceParagraphs?: readonly string[];
    readonly cacheParagraphs?: readonly string[];
    /** Whether either observed vector has non-whitespace characters; not proof of visibility. */
    readonly nonemptyText?: boolean;
  }[];
}
