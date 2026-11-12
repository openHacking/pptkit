import type { PaintInput, StrokeStyleInput } from "@pptkit/core";
import type { PptxTableStylePropertyInfo } from "../types/inheritance.js";
import type { PptxSourceLocation } from "../types/source.js";
import type { XmlElement } from "../xml/read-xml.js";
import { attribute, children, invalidStructure } from "../presentation/structure-xml.js";
import { contentReader, type ContentReader } from "./content-reader.js";
import type { ThemeTextResolution, ThemeTextScope } from "./theme-text-resolver.js";
import { transformTextColor } from "./color-transforms.js";

export type TableStyleDeclaration = Omit<PptxTableStylePropertyInfo, "property" | "target" | "rowIndex" | "columnIndex">;
export interface TableStyleData { part: string; root?: XmlElement; r: ContentReader; paths: WeakMap<XmlElement, string> }
type Placeholder = () => Promise<ThemeTextResolution>;
const colorNames = ["srgbClr", "schemeClr", "sysClr", "scrgbClr", "hslClr", "prstClr"];

/** Source paint/line interpretation; no package IO or Core/native identities escape. */
export function tableStyleProperties(namespace: string, checkpoint: () => Promise<void>, unmapped: (source: PptxSourceLocation, reason: string) => void) {
  const at = (data: TableStyleData, node: XmlElement, attr?: string): PptxSourceLocation => ({ part: data.part, path: `${data.paths.get(node)!}${attr ? `/@${attr}` : ""}` });
  const blocked = (data: TableStyleData, node: XmlElement, reason: string, region?: string): TableStyleDeclaration => { data.r.note(node, reason); return { source: at(data, node), status: "unsupported", reason, ...(region ? { region } : {}) }; };
  const color = async (node: XmlElement, data: TableStyleData, scope: ThemeTextScope, placeholder?: Placeholder): Promise<ThemeTextResolution> => {
    await checkpoint(); const r = data.r;
    if (node.namespaceUri !== namespace) return { status: "unsupported", reason: "Unsupported table color namespace." };
    const modifiers = children(node);
    let resolved: ThemeTextResolution;
    if (node.local === "schemeClr") {
      const token = r.attr(node, "val");
      if (token === "phClr") {
        const base = placeholder ? await placeholder() : undefined;
        if (!base || base.status === "unsupported") return { status: "unsupported", reason: base?.reason ?? "An unbound placeholder color blocks the declaration.", ...(base?.resolvedFrom ? { resolvedFrom: base.resolvedFrom } : {}) };
        if (!base.program) return { status: "unsupported", reason: "Placeholder color lacks a composable source program." };
        const program = { rgb: base.program.rgb, modifiers: [...base.program.modifiers, ...modifiers] };
        const result = transformTextColor(program.rgb, program.modifiers, namespace), resolvedFrom = [...base.resolvedFrom, at(data, node), ...modifiers.map(n => at(data, n))];
        resolved = result.status === "mapped" ? { ...result, resolvedFrom, program } : { ...result, resolvedFrom };
      } else resolved = await scope.color(data.part, node, data.paths);
    } else if (["srgbClr", "sysClr"].includes(node.local)) {
      const name = node.local === "sysClr" ? "lastClr" : "val", value = r.attr(node, name);
      if (node.local === "sysClr") r.attr(node, "val");
      if (!value && node.local === "sysClr") return { status: "unsupported", reason: "System table color lacks a portable lastClr fallback." };
      if (!value || !/^[\da-f]{6}$/i.test(value)) invalidStructure("Invalid table RGB color.", data.part);
      const result = transformTextColor(value!, modifiers, namespace), resolvedFrom = [at(data, node), ...modifiers.map(n => at(data, n))];
      resolved = result.status === "mapped" ? { ...result, resolvedFrom, program: { rgb: value!, modifiers } } : { ...result, resolvedFrom };
    } else return { status: "unsupported", reason: "Unsupported table color blocks an older declaration." };
    if (resolved.status === "mapped") for (const modifier of modifiers) { await checkpoint(); r.known(modifier); r.attr(modifier, "val"); }
    return resolved;
  };
  const paint = async (node: XmlElement, data: TableStyleData, scope: ThemeTextScope, region?: string, placeholder?: Placeholder): Promise<TableStyleDeclaration | undefined> => {
    await checkpoint(); const r = data.r;
    const fills = children(node).filter(n => n.namespaceUri === namespace && ["noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill", "fillRef"].includes(n.local));
    if (fills.length > 1) invalidStructure("Conflicting table fill choices.", data.part);
    const fill = fills[0]; if (!fill) return undefined;
    if (fill.local === "noFill") { r.known(fill); return { source: at(data, fill), status: "mapped", value: { type: "none" }, ...(region ? { region } : {}) }; }
    if (fill.local !== "solidFill") return blocked(data, fill, "Unsupported table fill blocks an older declaration.", region);
    r.known(fill); const choices = children(fill), selected = choices.length === 1 ? choices[0] : undefined;
    if (!selected) return blocked(data, fill, "Table solid fill requires one supported color.", region);
    const resolved = await color(selected, data, scope, placeholder);
    if (resolved.status === "unsupported") return { ...blocked(data, selected, resolved.reason, region), ...(resolved.resolvedFrom?.length ? { resolvedFrom: resolved.resolvedFrom } : {}) };
    return { source: at(data, selected), status: "mapped", value: { type: "solid", color: resolved.value, opacity: resolved.opacity ?? 1 }, ...(region ? { region } : {}), ...(resolved.resolvedFrom.length ? { resolvedFrom: resolved.resolvedFrom } : {}) };
  };
  const stroke = async (parent: XmlElement, name: string, data: TableStyleData, scope: ThemeTextScope, region?: string, placeholder?: Placeholder): Promise<TableStyleDeclaration | undefined> => {
    await checkpoint(); const r = data.r, line = r.one(parent, namespace, name); if (!line) return undefined;
    const value = r.stroke(parent, name)!, fill = await paint(line, data, scope, region, placeholder);
    for (const [key, neutral] of [["cap", "flat"], ["cmpd", "sng"], ["algn", "ctr"]] as const) if (attribute(line, key) === neutral) r.attr(line, key);
    const miter = r.one(line, namespace, "miter");
    let unsupportedJoin = false;
    if (miter) {
      r.known(miter); if (attribute(miter, "lim") !== undefined) r.number(miter, "lim", 0, 2147483647);
      unsupportedJoin = children(miter).length > 0 || miter.attributes.some(a => a.namespaceUri !== "http://www.w3.org/2000/xmlns/" && (a.namespaceUri !== "" || a.local !== "lim"));
    }
    const unsupported = unsupportedJoin || children(line).some(n => n.namespaceUri !== namespace || !["solidFill", "noFill", "prstDash", "headEnd", "tailEnd", "miter"].includes(n.local)) || line.attributes.some(a => a.namespaceUri !== "http://www.w3.org/2000/xmlns/" && (a.namespaceUri !== "" || !(a.local === "w" || a.local === "cmpd" && a.value === "sng" || a.local === "cap" && a.value === "flat" || a.local === "algn" && a.value === "ctr"))) || (value.beginArrow !== undefined && value.beginArrow !== "none") || (value.endArrow !== undefined && value.endArrow !== "none") || (r.one(line, namespace, "prstDash") !== undefined && value.dash === undefined);
    if (unsupported || !fill || fill.status === "unsupported") return { ...blocked(data, line, fill?.reason ?? "Unsupported or incomplete table line blocks an older border.", region), ...(fill?.resolvedFrom?.length ? { resolvedFrom: fill.resolvedFrom } : {}) };
    return { source: at(data, line), status: "mapped", value: { ...value, paint: fill.value as PaintInput }, ...(region ? { region } : {}), ...(fill.resolvedFrom?.length ? { resolvedFrom: fill.resolvedFrom } : {}) };
  };
  const matrixReference = async (ref: XmlElement, kind: "fill" | "line", data: TableStyleData, scope: ThemeTextScope, region?: string): Promise<TableStyleDeclaration> => {
    await checkpoint(); const r = data.r; r.known(ref);
    const index = r.number(ref, "idx", 0, 4294967295), source = at(data, ref, "idx");
    const choices = children(ref).filter(n => n.namespaceUri === namespace && colorNames.includes(n.local));
    if (choices.length > 1) invalidStructure("Conflicting style-reference colors.", data.part);
    if (index === 0 || kind === "fill" && index === 1000) return { source, status: "mapped", value: kind === "fill" ? { type: "none" } : { paint: { type: "none" }, width: 0 }, ...(region ? { region } : {}) };
    const entry = await scope.matrix(kind, index);
    if (entry.status === "unsupported") return { ...blocked(data, ref, entry.reason, region), ...(entry.resolvedFrom ? { resolvedFrom: entry.resolvedFrom } : {}) };
    let pending: Promise<ThemeTextResolution> | undefined;
    const placeholder: Placeholder = () => pending ??= choices[0] ? color(choices[0], data, scope) : Promise.resolve({ status: "unsupported", reason: "The selected style uses phClr but its reference has no supported color." });
    // Theme nodes stay in their original tree: wrapping never rewrites provenance.
    const selectedData: TableStyleData = { part: entry.part, paths: entry.paths, r: contentReader(entry.part, r.p, namespace) };
    const holder: XmlElement = { ...entry.node, children: [entry.node] };
    const selected = kind === "fill" ? await paint(holder, selectedData, scope, region, placeholder) : await stroke(holder, "ln", selectedData, scope, region, placeholder);
    const resolvedFrom = [...entry.resolvedFrom, ...(selected?.resolvedFrom ?? [])];
    await selectedData.r.audit(entry.node, entry.paths.get(entry.node)!, (path, reason) => unmapped({ part: entry.part, path }, `Theme matrix entry: ${reason}`), checkpoint);
    if (!selected || selected.status === "unsupported") return { ...blocked(data, ref, selected?.reason ?? "The selected matrix style is unsupported.", region), resolvedFrom };
    return { ...selected, source, resolvedFrom };
  };
  return { at, blocked, paint, stroke, matrixReference };
}
