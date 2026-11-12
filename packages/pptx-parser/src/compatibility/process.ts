import { PackageInspectionError } from "../types/archive.js";
import type { MarkupCompatibilityDecision, MarkupCompatibilityReport } from "../types/compatibility.js";
import type { XmlExpandedName } from "../types/xml.js";
import type { XmlAttribute, XmlElement, XmlNode } from "../xml/read-xml.js";
import { XML_ID } from "../xml/names.js";
import { expandedKey, MCE, XML, XMLNS, type CompatibilityConfiguration } from "./configuration.js";

const named = (node: XmlExpandedName): XmlExpandedName => ({ name: node.name, prefix: node.prefix, local: node.local, namespaceUri: node.namespaceUri });
const tokens = (value: string): string[] => value.split(/[\t\r\n ]+/).filter(Boolean);
interface Metadata {
  index: number;
  opaque: boolean;
  ignored: boolean;
  unwrapped: boolean;
  must: string[];
  requires: string[];
  ignoredAttributes: Set<XmlAttribute>;
}
type ScopeMap = Map<string, string>;

/** Private semantic projection; source trees are never mutated or serialized for preservation. */
export async function processCompatibility(root: XmlElement, configuration: CompatibilityConfiguration, part?: string, signal?: AbortSignal, interpretExtension?: (node: XmlElement, applicationParent?: XmlElement, applicationGrandparent?: XmlElement) => boolean): Promise<{ nodes: XmlNode[]; report: MarkupCompatibilityReport; origins: WeakMap<XmlElement, XmlElement> }> {
  const { understood, extensions } = configuration;
  const metadata = new Map<XmlElement, Metadata>();
  const namespaces: ScopeMap = new Map([["xml", XML]]), ignorable: ScopeMap = new Map(), process: ScopeMap = new Map();
  const invalid = (message: string): never => { throw new PackageInspectionError("mce-invalid", message, part); };
  const mismatch = (message: string): never => { throw new PackageInspectionError("mce-unsupported", message, part); };
  const abort = (): void => { if (signal?.aborted) throw new PackageInspectionError("operation-aborted", "MCE processing was aborted.", part); };
  let work = 0, index = 0;
  const checkpoint = async (): Promise<void> => { abort(); if (++work % 512 === 0) { await new Promise<void>((resolve) => setTimeout(resolve, 0)); abort(); } };
  const resolvePrefix = (prefix: string): string => {
    const uri = namespaces.get(prefix);
    if (!XML_ID.test(prefix) || !uri || uri === MCE) invalid("MCE prefix is unbound, invalid or bound to the MCE namespace.");
    return uri!;
  };
  // Scoped maps use reversible writes; no per-node copies of inherited namespace sets.
  type Undo = [ScopeMap, string, string | undefined];
  type Scan = { node: XmlElement; parent: XmlElement | undefined; applicationParent?: XmlElement; applicationGrandparent?: XmlElement; opaque: boolean } | { undo: Undo[] };
  const scanning: Scan[] = [{ node: root, parent: undefined, opaque: false }];
  while (scanning.length) {
    await checkpoint();
    const frame = scanning.pop()!;
    if ("undo" in frame) {
      for (let i = frame.undo.length - 1; i >= 0; i--) { const [map, key, before] = frame.undo[i]!; if (before === undefined) map.delete(key); else map.set(key, before); }
      continue;
    }
    // Only format-specific readers may opt a known extension back into processing.
    // An opaque ancestor always wins; the public generic API has no such override.
    const node = frame.node, opaque = frame.opaque || (extensions.has(expandedKey(node.namespaceUri, node.local)) && !interpretExtension?.(node, frame.applicationParent, frame.applicationGrandparent));
    const info: Metadata = { index: ++index, opaque, ignored: false, unwrapped: false, must: [], requires: [], ignoredAttributes: new Set() };
    metadata.set(node, info);
    const undo: Undo[] = [];
    const set = (map: ScopeMap, key: string, value: string): void => { undo.push([map, key, map.get(key)]); map.set(key, value); };
    if (!opaque) {
      for (const attribute of node.attributes) if (attribute.namespaceUri === XMLNS) set(namespaces, attribute.prefix === "xmlns" ? attribute.local : "", attribute.value);
      const mc = new Map(node.attributes.filter((a) => a.namespaceUri === MCE).map((a) => [a.local, a.value]));
      for (const name of mc.keys()) if (!["Ignorable", "ProcessContent", "MustUnderstand"].includes(name)) mismatch(`Unsupported MCE attribute ${name}.`);
      for (const prefix of tokens(mc.get("Ignorable") ?? "")) set(ignorable, resolvePrefix(prefix), "1");
      for (const token of tokens(mc.get("ProcessContent") ?? "")) {
        const pieces = token.split(":"), prefix = pieces[0]!, local = pieces[1];
        if (pieces.length !== 2 || !local || (local !== "*" && !XML_ID.test(local))) invalid("Invalid ProcessContent qualified name.");
        const uri = resolvePrefix(prefix);
        if (!ignorable.has(uri)) invalid("ProcessContent namespace must be ignorable in the current scope.");
        set(process, expandedKey(uri, local!), "1");
      }
      info.must = tokens(mc.get("MustUnderstand") ?? "").map(resolvePrefix);
      if (frame.parent?.namespaceUri === MCE && frame.parent.local === "AlternateContent" && node.namespaceUri !== MCE && !ignorable.has(node.namespaceUri)) invalid("Foreign AlternateContent child must have an ignorable namespace.");
      const unknownIgnorable = ignorable.has(node.namespaceUri) && !understood.has(node.namespaceUri);
      info.unwrapped = unknownIgnorable && (process.has(expandedKey(node.namespaceUri, node.local)) || process.has(expandedKey(node.namespaceUri, "*")));
      info.ignored = unknownIgnorable && !info.unwrapped;
      for (const a of node.attributes) {
        if (ignorable.has(a.namespaceUri) && !understood.has(a.namespaceUri)) info.ignoredAttributes.add(a);
        if (info.unwrapped && a.namespaceUri === XML && ["base", "lang", "space"].includes(a.local)) invalid("Unwrapped elements cannot carry xml:base, xml:lang or xml:space.");
      }
      if (node.namespaceUri === MCE) {
        if (!["AlternateContent", "Choice", "Fallback"].includes(node.local)) mismatch(`Unsupported MCE element ${node.local}.`);
        if (node.local !== "AlternateContent" && (frame.parent?.namespaceUri !== MCE || frame.parent.local !== "AlternateContent")) invalid("Choice/Fallback must be a direct child of AlternateContent.");
        for (const a of node.attributes) {
          if (a.namespaceUri === XMLNS || a.namespaceUri === MCE) continue;
          if (node.local === "Choice" && !a.namespaceUri && a.local === "Requires") continue;
          if (!a.namespaceUri || a.namespaceUri === XML || !ignorable.has(a.namespaceUri)) invalid("Invalid attribute on MCE structural element.");
        }
        if (node.local === "Choice") {
          const requires = node.attributes.find((a) => !a.namespaceUri && a.local === "Requires");
          if (!requires || !tokens(requires.value).length) invalid("Choice requires a nonempty unqualified Requires attribute.");
          info.requires = tokens(requires!.value).map(resolvePrefix);
        }
        if (node.local === "AlternateContent") {
          let choices = 0, fallback = false;
          for (const child of node.children) {
            if (child.kind !== "element") {
              if ((child.kind === "text" || child.kind === "cdata") && tokens(child.value).length) invalid("AlternateContent cannot contain character content.");
              continue;
            }
            if (child.namespaceUri !== MCE) {
              // Validate in the child's scope, including its own Ignorable declaration.
              continue;
            }
            if (child.local === "Choice" && !fallback) choices++;
            else if (child.local === "Fallback" && !fallback) fallback = true;
            else invalid("Invalid AlternateContent child ordering or multiplicity.");
          }
          if (!choices) invalid("AlternateContent requires at least one Choice.");
        }
      }
    }
    scanning.push({ undo });
    const transparent = node.namespaceUri === MCE || info.unwrapped;
    const applicationParent = transparent ? frame.applicationParent : node, applicationGrandparent = transparent ? frame.applicationGrandparent : frame.applicationParent;
    for (let i = node.children.length - 1; i >= 0; i--) { const child = node.children[i]!; if (child.kind === "element") scanning.push({ node: child, parent: node, ...(applicationParent ? { applicationParent } : {}), ...(applicationGrandparent ? { applicationGrandparent } : {}), opaque }); }
  }
  const nodes: XmlNode[] = [], decisions: MarkupCompatibilityDecision[] = [], origins = new WeakMap<XmlElement, XmlElement>();
  type OutputFrame = { node: XmlNode; output: XmlNode[]; opaque?: boolean };
  const outputStack: OutputFrame[] = [{ node: root, output: nodes }];
  const requireUnderstood = (info: Metadata): void => { for (const uri of info.must) if (!understood.has(uri)) mismatch(`MustUnderstand namespace is not understood: ${uri}`); };
  const enqueue = (children: readonly XmlNode[], output: XmlNode[], opaque = false): void => { for (let i = children.length - 1; i >= 0; i--) outputStack.push({ node: children[i]!, output, opaque }); };
  while (outputStack.length) {
    await checkpoint();
    const { node, output, opaque } = outputStack.pop()!;
    if (node.kind !== "element") { output.push({ ...node }); continue; }
    const info = metadata.get(node)!;
    const decision = (action: MarkupCompatibilityDecision["action"], extra = {}): void => { decisions.push({ elementIndex: info.index, name: named(node), action, ...extra }); };
    if (info.opaque || opaque) {
      const copy: XmlElement = { ...node, attributes: node.attributes.map((attribute) => ({ ...attribute })), children: [] };
      origins.set(copy, node);
      output.push(copy); enqueue(node.children, copy.children, true);
      if (!opaque) decision("opaque-extension");
      continue;
    }
    if (info.ignored) { decision("ignored-element"); continue; }
    requireUnderstood(info);
    if (info.unwrapped) { decision("unwrapped-element"); enqueue(node.children, output); continue; }
    if (node.namespaceUri === MCE) {
      if (node.local !== "AlternateContent") invalid("Unexpected standalone MCE branch during projection.");
      let selected: XmlElement | undefined;
      for (const child of node.children) if (child.kind === "element") {
        const childInfo = metadata.get(child)!;
        if (child.namespaceUri !== MCE) { if (!childInfo.ignored) mismatch("AlternateContent has a non-ignored foreign child."); continue; }
        if (!selected && (child.local === "Fallback" || childInfo.requires.every((uri) => understood.has(uri)))) selected = child;
      }
      decision("alternate-content", selected ? { selectedElementIndex: metadata.get(selected)!.index } : {});
      if (selected) { requireUnderstood(metadata.get(selected)!); enqueue(selected.children, output); }
      continue;
    }
    if (!understood.has(node.namespaceUri)) mismatch(`Element namespace is not understood: ${node.namespaceUri}`);
    const attributes: XmlAttribute[] = [];
    for (const attribute of node.attributes) {
      if (attribute.namespaceUri === MCE) continue;
      if (info.ignoredAttributes.has(attribute)) { decision("ignored-attribute", { attribute: named(attribute) }); continue; }
      if (!understood.has(attribute.namespaceUri)) mismatch(`Attribute namespace is not understood: ${attribute.namespaceUri}`);
      attributes.push({ ...attribute });
    }
    const copy: XmlElement = { ...node, attributes, children: [] };
    origins.set(copy, node);
    output.push(copy); enqueue(node.children, copy.children);
  }
  const summary = { roots: nodes.filter((node): node is XmlElement => node.kind === "element").map(named), elements: 0, attributes: 0, textCharacters: 0 };
  const counting = [...nodes];
  while (counting.length) {
    await checkpoint();
    const node = counting.pop()!;
    if (node.kind === "element") { summary.elements++; summary.attributes += node.attributes.length; for (const child of node.children) counting.push(child); }
    else if (node.kind === "text" || node.kind === "cdata") summary.textCharacters += node.value.length;
  }
  abort();
  return { nodes, report: { output: summary, decisions }, origins };
}
