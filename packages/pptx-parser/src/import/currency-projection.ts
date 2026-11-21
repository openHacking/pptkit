import type { CustomDocumentProperty } from "@pptkit/core";
import { XMLNS } from "../compatibility/configuration.js";
import type { structureContext } from "../presentation/structure-context.js";
import { children } from "../presentation/structure-xml.js";
import { projectCurrencyProperty } from "./project-custom-properties.js";

/** Project effective direct currency text; lexical text remains private source evidence. */
export async function projectCurrencies(context: Awaited<ReturnType<typeof structureContext>>, checkpoint: () => Promise<void>): Promise<{ values: (CustomDocumentProperty | undefined)[]; lexical: (string | undefined)[] }> {
  const values: (CustomDocumentProperty | undefined)[] = [], lexical: (string | undefined)[] = [];
  if (!context.metadata.custom || !context.customMetadataRoot) return { values, lexical };
  const nodes = children(context.customMetadataRoot);
  for (const [index, property] of context.metadata.custom.properties.entries()) {
    await checkpoint();
    const node = nodes[index], value = node && children(node)[0];
    if (property.valueType !== "cy" || !value || value.attributes.some(attr => attr.namespaceUri !== XMLNS)) { values.push(undefined); lexical.push(undefined); continue; }
    let text = "", invalid = false;
    for (const child of value.children) {
      await checkpoint();
      if ((child.kind !== "text" && child.kind !== "cdata") || text.length + child.value.length > 1000000) { invalid = true; break; }
      text += child.value;
    }
    const projected = invalid ? undefined : projectCurrencyProperty(property, text);
    values.push(projected); lexical.push(projected ? text : undefined);
  }
  return { values, lexical };
}
