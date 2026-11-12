import type { XmlInspectionOptions } from "../types/xml.js";
import { appendPropertyDeclarations } from "./append-property-declarations.js";

/** Compatibility wrapper for the custom-property declaration path. */
export function appendCustomPropertyDeclarations(relationships: Uint8Array, contentTypes: Uint8Array, request: { readonly part: string; readonly relationshipId: string; readonly dialect: "strict" | "transitional" }, options: XmlInspectionOptions & { readonly maxOutputBytes: number }): Promise<{ relationships: Uint8Array; contentTypes: Uint8Array }> {
  if (!request) return appendPropertyDeclarations(relationships, contentTypes, request as never, options);
  const { part, relationshipId, dialect } = request;
  return appendPropertyDeclarations(relationships, contentTypes, { part, relationshipId, dialect, kind: "custom" }, options);
}
