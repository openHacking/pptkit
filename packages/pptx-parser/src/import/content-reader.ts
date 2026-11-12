import { attribute, children, integer, invalidStructure, singleChild } from "../presentation/structure-xml.js";
import type { XmlElement } from "../xml/read-xml.js";
import type { PaintInput, StrokeStyleInput, TransformInput, Box } from "@pptkit/core";

/** Tracks exactly which source nodes/attributes were interpreted. */
export function contentReader(part: string, p: string, a: string) {
  const consumed = new WeakMap<XmlElement, Set<string>>();
  const textLeaves = new WeakSet<XmlElement>();
  const notes = new WeakMap<XmlElement, string[]>();
  const note = (node: XmlElement, reason: string) => { known(node); const values = notes.get(node) ?? []; values.push(reason); notes.set(node, values); };
  const known = (node: XmlElement) => { if (!consumed.has(node)) consumed.set(node, new Set()); return node; };
  const attr = (node: XmlElement, name: string, namespaceUri = "") => { known(node); consumed.get(node)!.add(`{${namespaceUri}}${name}`); return attribute(node, name, namespaceUri); };
  const one = (node: XmlElement, ns: string, name: string, required = false) => singleChild(node, ns, name, part, required);
  const number = (node: XmlElement, name: string, min: number, max: number, fallback?: number) => {
    const value = attr(node, name);
    return value === undefined && fallback !== undefined ? fallback : integer(value, min, max, `${node.local}/@${name}`, part);
  };
  const bool = (node: XmlElement, name: string, fallback = false) => {
    const value = attr(node, name)?.trim();
    if (value === undefined) return fallback;
    if (value === "1" || value === "true") return true;
    if (value === "0" || value === "false") return false;
    return invalidStructure(`Invalid boolean ${node.local}/@${name}.`, part);
  };
  const enumValue = <T extends string>(node: XmlElement, name: string, mapping: Readonly<Record<string, T>>): T | undefined => {
    const value = attribute(node, name);
    if (value !== undefined && Object.hasOwn(mapping, value)) { attr(node, name); return mapping[value]; }
    return undefined;
  };
  const color = (node: XmlElement): { color: string; opacity: number } | undefined => {
    const values = children(node);
    if (values.length !== 1 || values[0]!.namespaceUri !== a || values[0]!.local !== "srgbClr") return undefined;
    const rgb = values[0]!, value = attribute(rgb, "val");
    if (!value || !/^[\dA-Fa-f]{6}$/.test(value)) invalidStructure("Invalid sRGB color.", part);
    known(node); attr(rgb, "val");
    const alpha = one(rgb, a, "alpha");
    const opacity = alpha ? number(alpha, "val", 0, 100000) / 100000 : 1;
    return { color: value!.toUpperCase(), opacity };
  };
  const paint = (node: XmlElement): PaintInput | undefined => {
    const fills = children(node).filter(n => n.namespaceUri === a && ["noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"].includes(n.local));
    if (fills.length > 1) invalidStructure("Conflicting drawing fill choices.", part);
    const fill = fills[0];
    if (fill?.local === "noFill") { known(fill); return { type: "none" }; }
    if (fill?.local === "solidFill") { const rgb = color(fill); if (rgb) return { type: "solid", ...rgb }; }
    return undefined;
  };
  const stroke = (node: XmlElement, name = "ln"): StrokeStyleInput | undefined => {
    const line = one(node, a, name);
    if (!line) return undefined;
    known(line);
    const linePaint = paint(line), width = number(line, "w", 0, 20116800, 12700) / 12700;
    const dashNode = one(line, a, "prstDash");
    const dash = dashNode ? enumValue(dashNode, "val", { solid: "solid", dash: "dash", dot: "dot", sysDot: "dot", dashDot: "dashDot" } as const) : undefined;
    const arrow = (name: string) => {
      const child = one(line, a, name);
      return child ? enumValue(child, "type", { none: "none", arrow: "arrow", triangle: "triangle", stealth: "stealth", diamond: "diamond", oval: "oval" } as const) : undefined;
    };
    const beginArrow = arrow("headEnd"), endArrow = arrow("tailEnd");
    return { width, ...(linePaint ? { paint: linePaint } : {}), ...(dash ? { dash } : {}), ...(beginArrow ? { beginArrow } : {}), ...(endArrow ? { endArrow } : {}) };
  };
  const transform = (node: XmlElement, namespace = a, local = "xfrm"): { box: Box; transform: TransformInput } | undefined => {
    const xfrm = one(node, namespace, local);
    if (!xfrm) return undefined;
    known(xfrm);
    const off = one(xfrm, a, "off"), ext = one(xfrm, a, "ext");
    if (!off || !ext) return undefined;
    const box = { x: number(off, "x", -27273042329600, 27273042316900) / 12700, y: number(off, "y", -27273042329600, 27273042316900) / 12700, width: number(ext, "cx", 0, 27273042316900) / 12700, height: number(ext, "cy", 0, 27273042316900) / 12700 };
    return { box, transform: { rotation: number(xfrm, "rot", -2147483648, 2147483647, 0) / 60000, flipH: bool(xfrm, "flipH"), flipV: bool(xfrm, "flipV") } };
  };
  const text = (node: XmlElement): string => {
    if (children(node).length) invalidStructure("A scalar text leaf cannot contain child elements.", part);
    known(node); textLeaves.add(node);
    return node.children.map(n => n.kind === "text" || n.kind === "cdata" ? n.value : "").join("");
  };
  return { part, p, a, known, attr, one, number, bool, enumValue, color, paint, stroke, transform, note, text,
    async audit(root: XmlElement, path: string, unhandled: (path: string, reason: string) => void, checkpoint: () => Promise<void>) {
      const stack = [{ node: root, path }];
      while (stack.length) {
        await checkpoint();
        const item = stack.pop()!, used = consumed.get(item.node);
        if (!used) { unhandled(item.path, "Source subtree is not interpreted by the content profile."); continue; }
        for (const reason of notes.get(item.node) ?? []) unhandled(item.path, reason);
        for (const at of item.node.attributes) if (at.namespaceUri !== "http://www.w3.org/2000/xmlns/" && !used.has(`{${at.namespaceUri}}${at.local}`)) unhandled(`${item.path}/@${at.name}`, "Source attribute is not mapped to Core.");
        if (!textLeaves.has(item.node) && !(item.node.namespaceUri === a && item.node.local === "t") && item.node.children.some(n => (n.kind === "text" || n.kind === "cdata") && !/^[\t\r\n ]*$/.test(n.value))) unhandled(`${item.path}/text()`, "Character content outside a mapped text leaf is not interpreted.");
        const counts = new Map<string, number>(), nodes = children(item.node);
        for (const child of nodes) {
          const key = `{${child.namespaceUri}}${child.local}`, index = (counts.get(key) ?? 0) + 1;
          counts.set(key, index);
          stack.push({ node: child, path: `${item.path}/${child.name}[${index}]` });
        }
      }
    },
  };
}
export type ContentReader = ReturnType<typeof contentReader>;

/** Effective source paths preserve prefixes while counting by expanded name. */
export async function contentPaths(root: XmlElement, checkpoint: () => Promise<void>): Promise<WeakMap<XmlElement, string>> {
  const paths = new WeakMap<XmlElement, string>(), stack = [{ node: root, path: `/${root.name}[1]` }];
  while (stack.length) {
    await checkpoint();
    const item = stack.pop()!; paths.set(item.node, item.path);
    const counts = new Map<string, number>();
    for (const node of children(item.node)) {
      const key = `{${node.namespaceUri}}${node.local}`, position = (counts.get(key) ?? 0) + 1;
      counts.set(key, position);
      stack.push({ node, path: `${item.path}/${node.name}[${position}]` });
    }
  }
  return paths;
}
