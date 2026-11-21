import { snapshotOptions } from "../opc/read-package.js";
import { structureContext } from "../presentation/structure-context.js";
import { children } from "../presentation/structure-xml.js";
import { PackageInspectionError } from "../types/archive.js";
import type { OpcInspectionOptions } from "../types/opc.js";
import { byteViewInfo } from "../utils/byte-view.js";

export interface OpaqueCustomPropertyValue {
  readonly part: string;
  readonly id: number;
  readonly formatId: string;
  readonly name?: string;
  readonly linkTarget?: string;
  readonly valueType: string;
  readonly valueNamespace: string;
  readonly content: { readonly status: "text"; readonly value: string } | { readonly status: "non-text" };
}

/** Inspect effective, still-opaque custom values; never coerce, dereference or grant edit access. */
export async function inspectOpaqueCustomProperties(bytes: Uint8Array, options: OpcInspectionOptions = {}): Promise<readonly OpaqueCustomPropertyValue[]> {
  if (!(bytes instanceof Uint8Array)) throw new TypeError("PPTX input must be Uint8Array.");
  if (byteViewInfo(bytes).shared) throw new TypeError("SharedArrayBuffer input is not supported.");
  const captured = snapshotOptions(options);
  const context = await structureContext(bytes, { limits: captured.limits, xmlLimits: captured.xmlLimits, zipLimits: captured.zipLimits, ...(captured.signal ? { signal: captured.signal } : {}) });
  const metadata = context.metadata.custom, root = context.customMetadataRoot;
  if (!metadata || !root) return Object.freeze([]);
  const nodes = children(root), output: OpaqueCustomPropertyValue[] = [];
  let textCharacters = 0;
  for (let i = 0; i < metadata.properties.length; i++) {
    await context.checkpoint();
    const property = metadata.properties[i]!;
    if (property.value.status !== "uninterpreted") continue;
    if (output.length >= 100000) throw new PackageInspectionError("package-limit-exceeded", "Opaque custom property record limit exceeded.", metadata.part);
    const value = children(nodes[i]!)[0]!;
    let content: OpaqueCustomPropertyValue["content"];
    let textOnly = true;
    for (const child of value.children) {
      await context.checkpoint();
      if (child.kind !== "text" && child.kind !== "cdata") { textOnly = false; break; }
    }
    if (textOnly) {
      const segments: string[] = [];
      for (const child of value.children) {
        await context.checkpoint();
        if (child.kind !== "text" && child.kind !== "cdata") continue;
        textCharacters += child.value.length;
        if (textCharacters > 1000000) throw new PackageInspectionError("package-limit-exceeded", "Opaque custom property text limit exceeded.", metadata.part);
        segments.push(child.value);
      }
      content = Object.freeze({ status: "text", value: segments.join("") });
    } else content = Object.freeze({ status: "non-text" });
    output.push(Object.freeze({ part: metadata.part, id: property.id, formatId: property.formatId, ...(property.name === undefined ? {} : { name: property.name }), ...(property.linkTarget === undefined ? {} : { linkTarget: property.linkTarget }), valueType: property.valueType, valueNamespace: property.value.namespaceUri, content }));
  }
  return Object.freeze(output);
}
