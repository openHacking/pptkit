import type { CustomPropertyInfo, ExtendedPropertyInfo } from "../types/metadata.js";
import type { XmlElement } from "../xml/read-xml.js";
import { attribute, children, invalidStructure } from "../presentation/structure-xml.js";
import { attributes, integerValue, noText, textValue, trimmed } from "./xml.js";
import { scalar, uninterpreted, variant } from "./values.js";

const STRINGS = new Set(["Template", "Manager", "Company", "PresentationFormat", "HyperlinkBase", "Application", "AppVersion"]);
const INTEGERS = new Set(["Pages", "Words", "Characters", "Lines", "Paragraphs", "Slides", "Notes", "TotalTime", "HiddenSlides", "MMClips", "CharactersWithSpaces", "DocSecurity"]);
const BOOLEANS = new Set(["ScaleCrop", "LinksUpToDate", "SharedDoc", "HyperlinksChanged"]);
const COMPLEX = new Set(["HeadingPairs", "TitlesOfParts", "HLinks", "DigSig"]);
export function extendedProperties(root: XmlElement, vt: string, part: string): ExtendedPropertyInfo[] {
  attributes(root, [], part); noText(root, part);
  const result: ExtendedPropertyInfo[] = [], seen = new Set<string>();
  for (const node of children(root)) {
    if (node.namespaceUri !== root.namespaceUri || seen.has(node.local)) invalidStructure("Unexpected or duplicate extended property.", part);
    seen.add(node.local);
    if (COMPLEX.has(node.local)) {
      attributes(node, [], part); noText(node, part);
      const values = children(node), local = node.local === "DigSig" ? "blob" : "vector";
      if (values.length !== 1 || values[0]!.local !== local || values[0]!.namespaceUri !== vt) invalidStructure(`Expected one vt:${local} value.`, part);
      result.push({ name: node.local, value: uninterpreted(values[0]!) });
    } else {
      const kind = STRINGS.has(node.local) ? "string" : INTEGERS.has(node.local) ? "integer" : BOOLEANS.has(node.local) ? "boolean" : undefined;
      if (!kind) invalidStructure(`Unknown extended property ${node.local}.`, part);
      const value = scalar(node, kind!, part);
      if (kind === "integer") integerValue(textValue(node, part), -2147483648n, 2147483647n, part);
      result.push({ name: node.local, value });
    }
  }
  return result;
}
export async function customProperties(root: XmlElement, vt: string, part: string, checkpoint: () => Promise<void>): Promise<CustomPropertyInfo[]> {
  attributes(root, [], part); noText(root, part);
  const result: CustomPropertyInfo[] = [], identities = new Set<string>(), names = new Set<string>();
  for (const node of children(root)) {
    await checkpoint();
    if (node.namespaceUri !== root.namespaceUri || node.local !== "property") invalidStructure("Unexpected custom property element.", part);
    attributes(node, ["|pid", "|fmtid", "|name", "|linkTarget"], part); noText(node, part);
    const pid = attribute(node, "pid") ?? "", formatId = trimmed(attribute(node, "fmtid") ?? "");
    integerValue(pid, -2147483648n, 2147483647n, part);
    if (!/^\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}$/.test(formatId)) invalidStructure("Invalid custom property format GUID.", part);
    const id = Number(trimmed(pid)), name = attribute(node, "name"), linkTarget = attribute(node, "linkTarget"), identity = `${formatId}:${id}`;
    if (identities.has(identity) || (name !== undefined && names.has(name))) invalidStructure("Duplicate custom property identity/name.", part);
    identities.add(identity); if (name !== undefined) names.add(name);
    const values = children(node);
    if (values.length !== 1) invalidStructure("Custom property must contain one typed value.", part);
    const value = variant(values[0]!, vt, part);
    result.push({ id, formatId, ...(name !== undefined ? { name } : {}), ...(linkTarget !== undefined ? { linkTarget } : {}), valueType: values[0]!.local, value });
  }
  return result;
}
