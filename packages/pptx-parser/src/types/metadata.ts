export interface MetadataKeywordSegment {
  readonly kind: "text" | "value";
  readonly value: string;
  readonly language?: string;
}
export interface CorePropertyInfo {
  readonly name: "category" | "contentStatus" | "created" | "creator" | "description" | "identifier" | "keywords" | "language" | "lastModifiedBy" | "lastPrinted" | "modified" | "revision" | "subject" | "title" | "version";
  /** XML-normalized text, not trimmed or converted into a Date/number. */
  readonly value: string;
  readonly language?: string;
  readonly keywordSegments?: readonly MetadataKeywordSegment[];
}
export type MetadataScalarKind = "string" | "integer" | "boolean" | "date" | "real" | "decimal" | "empty" | "null";
export type MetadataPropertyValue = { readonly status: "scalar"; readonly kind: MetadataScalarKind; readonly value: string } | { readonly status: "uninterpreted"; readonly namespaceUri: string; readonly local: string };
export interface ExtendedPropertyInfo {
  readonly name: string;
  readonly value: MetadataPropertyValue;
}
export interface CustomPropertyInfo {
  readonly id: number;
  readonly formatId: string;
  readonly name?: string;
  readonly linkTarget?: string;
  readonly valueType: string;
  readonly value: MetadataPropertyValue;
}
export interface MetadataPartInfo<T> {
  readonly part: string;
  readonly relationshipId: string;
  /** Source child order, including empty values; property identities are not object keys. */
  readonly properties: readonly T[];
}
export interface PptxDocumentMetadata {
  readonly core?: MetadataPartInfo<CorePropertyInfo>;
  readonly extended?: MetadataPartInfo<ExtendedPropertyInfo>;
  readonly custom?: MetadataPartInfo<CustomPropertyInfo>;
}
