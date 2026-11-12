import { readParallelogramSlant } from "./parallelogram-slant.js";
import type { Box, GroupElementInput, PresentationElementInput, ShapeKind } from "@pptkit/core";
import type { structureContext } from "../presentation/structure-context.js";
import type { PptxDiagramInfo } from "../types/smartart.js";
import type { PptxSourceLocation } from "../types/source.js";
import type { XmlElement } from "../xml/read-xml.js";
import { children } from "../presentation/structure-xml.js";
import { contentReader, contentPaths } from "./content-reader.js";
import { contentText } from "./content-text.js";
import type { contentImages } from "./content-images.js";

const DSP = "http://schemas.microsoft.com/office/drawing/2008/diagram";
const MCE = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const presets: Readonly<Record<string, ShapeKind>> = { rect: "rect", roundRect: "roundRect", ellipse: "ellipse", triangle: "triangle", diamond: "diamond", rightArrow: "arrow", chevron: "chevron", parallelogram: "parallelogram" };
type Context = Awaited<ReturnType<typeof structureContext>>;
type PreviewBase = Pick<GroupElementInput, "id" | "name" | "hidden" | "transform" | "accessibility"> & { box: Box };

function placeholder(id: string, box: Box, label = true): PresentationElementInput {
  return { type: "shape", shape: "rect", id, name: "SmartArt", box,
    style: { fill: { type: "solid", color: "F2F2F2" }, stroke: { paint: { type: "solid", color: "777777" }, width: 1, dash: "dash" } },
    ...(label ? { text: { content: "SmartArt", frame: { verticalAlign: "middle" as const } } } : {}) };
}

/** Cached shapes are an explicit visual fallback. No data-node or source-edit bindings. */
export async function diagramPreview(
  diagram: PptxDiagramInfo | undefined, base: PreviewBase, context: Context,
  image: ReturnType<typeof contentImages>, checkpoint: () => Promise<void>,
  note: (source: PptxSourceLocation, reason: string) => void,
): Promise<{ input: GroupElementInput; status: "cache" | "partial-cache" | "placeholder"; cachePart?: string; reason?: string }> {
  const id = base.id!, extent = { width: base.box.width, height: base.box.height };
  const empty = (reason: string, cachePart?: string) => ({ input: { ...base, type: "group" as const, coordinateSize: { width: Math.max(1, extent.width), height: Math.max(1, extent.height) },
    children: [placeholder(`${id}-placeholder`, { x: 0, y: 0, ...extent })] }, status: "placeholder" as const, reason, ...(cachePart ? { cachePart } : {}) });
  const cache = diagram?.cache;
  if (cache?.status !== "located" || !cache.part) return empty(`No located cache: ${cache?.status ?? "unavailable"}.`);
  const part = cache.part, inventory = context.entry.package.parts.find(p => p.name === part);
  if (inventory?.xml?.namespaceUris.includes(MCE)) return empty("Cached drawing MCE is outside the preview profile.", part);
  if (!extent.width || !extent.height) return empty("A nonzero diagram frame is required for cached preview.", part);
  const root = await context.root(part, "drawing", undefined, DSP), r = contentReader(part, DSP, context.a);
  const paths = await contentPaths(root, checkpoint);
  let partial = false;
  const warn = (node: XmlElement, reason: string) => { partial = true; note({ part, path: paths.get(node)! }, reason); };
  const tree = r.one(root, DSP, "spTree"); r.known(root);
  if (!tree) return empty("Cached drawing has no shape tree.", part);
  const visual = children(tree).filter(n => !(n.namespaceUri === DSP && ["nvGrpSpPr", "grpSpPr"].includes(n.local)));
  if (!visual.length) return empty("Cached drawing shape tree is empty.", part);
  const meta = (node: XmlElement, group: boolean) => {
    const nv = r.one(node, DSP, group ? "nvGrpSpPr" : "nvSpPr"), props = nv ? r.one(nv, DSP, "cNvPr") : undefined;
    if (nv) {
      r.known(nv); const flags = r.one(nv, DSP, group ? "cNvGrpSpPr" : "cNvSpPr"); if (flags) r.known(flags);
    }
    r.attr(node, "modelId");
    if (!props) return {};
    r.attr(props, "id");
    return { name: r.attr(props, "name") ?? "", hidden: r.bool(props, "hidden") };
  };
  const groupFields = (props: XmlElement, size: { width: number; height: number }) => {
    const xfrm = r.one(props, r.a, "xfrm");
    if (!xfrm) return { origin: { x: 0, y: 0 }, coordinateSize: size };
    const off = r.one(xfrm, r.a, "chOff"), ext = r.one(xfrm, r.a, "chExt");
    if (!off || !ext) return undefined;
    const width = r.number(ext, "cx", 1, 27273042316900) / 12700, height = r.number(ext, "cy", 1, 27273042316900) / 12700;
    return { origin: { x: r.number(off, "x", -27273042329600, 27273042316900) / 12700, y: r.number(off, "y", -27273042329600, 27273042316900) / 12700 }, coordinateSize: { width, height } };
  };
  const parse = async (node: XmlElement, origin: { x: number; y: number }, parentSize: { width: number; height: number }): Promise<PresentationElementInput> => {
    await checkpoint();
    const index = (await context.sourceElement(part, node)).elementIndex, childId = `${id}-cache-${index}`;
    const group = node.namespaceUri === DSP && node.local === "grpSp";
    const props = r.one(node, DSP, group ? "grpSpPr" : "spPr"), geometry = props ? r.transform(props) : undefined;
    if (node.namespaceUri !== DSP || !["sp", "grpSp"].includes(node.local) || !props || !geometry) {
      warn(node, "Cache object or local geometry is unsupported; a parent-sized placeholder marks its unavailable position.");
      return { ...placeholder(childId, { x: 0, y: 0, ...parentSize }), ...(node.namespaceUri === DSP && ["sp", "grpSp"].includes(node.local) ? meta(node, group) : {}) };
    }
    r.known(node); r.known(props);
    geometry.box.x -= origin.x; geometry.box.y -= origin.y;
    const common = { id: childId, ...meta(node, group), ...geometry };
    if (group) {
      const fields = groupFields(props, geometry.box);
      if (!fields) { warn(node, "Cache group lacks a complete child coordinate space."); return { ...placeholder(childId, geometry.box), ...common }; }
      const result: PresentationElementInput[] = [];
      for (const child of children(node)) if (!(child.namespaceUri === DSP && ["nvGrpSpPr", "grpSpPr"].includes(child.local))) result.push(await parse(child, fields.origin, fields.coordinateSize));
      return { ...common, type: "group", coordinateSize: fields.coordinateSize, children: result };
    }
    const preset = r.one(props, r.a, "prstGeom"), custom = r.one(props, r.a, "custGeom"), kind = preset ? r.attr(preset, "prst") : undefined;
    const slant = kind === "parallelogram" && preset ? readParallelogramSlant(preset, r.a) : undefined;
    const tx = r.one(node, DSP, "txXfrm"), supported = kind !== undefined && Object.hasOwn(presets, kind) && !custom && (!slant || slant.supported);
    if (supported) { const adjustments = r.one(preset!, r.a, "avLst"); if (adjustments) r.known(adjustments); if (slant?.supported && slant.guide) { r.known(slant.guide); r.attr(slant.guide, "name"); r.attr(slant.guide, "fmla"); } }
    const body = r.one(node, DSP, "txBody");
    const textGeometry = tx && body ? r.transform(node, DSP, "txXfrm") : undefined;
    const bodyPr = tx && body ? r.one(body, r.a, "bodyPr") : undefined;
    const upright = bodyPr ? r.bool(bodyPr, "upright") : false;
    const bodyRotation = bodyPr ? r.number(bodyPr, "rot", -2147483648, 2147483647, 0) / 60000 : 0;
    const independentSupported = textGeometry && textGeometry.box.width > 0 && textGeometry.box.height > 0
      && !geometry.transform.flipH && !geometry.transform.flipV && !textGeometry.transform.flipH && !textGeometry.transform.flipV && !upright;
    const text = body && (tx ? independentSupported : supported) ? await contentText(body, r, checkpoint, () => undefined) : undefined;
    const visual = async (visualText: typeof text): Promise<PresentationElementInput> => {
      if (!supported) {
        warn(node, "Custom/preset geometry is outside the cache preview subset; separately located text can still be projected.");
        return { ...placeholder(childId, geometry.box, !tx), ...common };
      }
      const fill = r.one(props, r.a, "blipFill");
      if (fill) {
        if (!geometry.box.width || !geometry.box.height) { warn(node, "Zero image extent has no cache preview coordinate space."); return { ...placeholder(childId, geometry.box, !tx), ...common }; }
        const result = await image(node, props, r, { part, path: paths.get(node)! }, fill);
        if ("reason" in result) { warn(node, result.reason); return { ...placeholder(childId, geometry.box, !tx), ...common }; }
        const childBox = { x: 0, y: 0, width: geometry.box.width, height: geometry.box.height };
        return { ...common, type: "group", coordinateSize: { width: childBox.width, height: childBox.height }, children: [
          { id: `${childId}-image`, box: childBox, ...result.fields },
          ...(visualText ? [{ id: `${childId}-text`, type: "text" as const, box: childBox, ...visualText }] : []),
        ] };
      }
      return { ...common, type: "shape", shape: presets[kind!]!, ...(slant?.supported && slant.value !== undefined ? { slant: slant.value } : {}), style: { fill: r.paint(props) ?? { type: "none" }, stroke: r.stroke(props) ?? { paint: { type: "none" } } }, ...(visualText ? { text: visualText } : {}) };
    };
    if (!tx) return visual(text);
    const drawing = await visual(undefined); drawing.id = `${childId}-visual`;
    const independent: PresentationElementInput[] = [drawing];
    if (body) {
      if (!independentSupported || !textGeometry || !text) {
        warn(tx, "Independent text needs a complete nonzero rectangle without flips or upright semantics; text is marked unavailable.");
        independent.push(placeholder(`${childId}-text`, textGeometry ? { ...textGeometry.box, x: textGeometry.box.x - origin.x, y: textGeometry.box.y - origin.y } : geometry.box));
      } else {
        textGeometry.box.x -= origin.x; textGeometry.box.y -= origin.y;
        const rotation = ((geometry.transform.rotation! + textGeometry.transform.rotation! + bodyRotation) % 360 + 360) % 360;
        independent.push({ id: `${childId}-text`, type: "text", name: common.name ?? "", box: textGeometry.box,
          transform: { rotation, flipH: false, flipV: false }, ...text });
      }
    } else { warn(tx, "Independent text transform has no mapped text body."); }
    return { ...common, type: "group", box: { x: 0, y: 0, ...parentSize }, transform: { rotation: 0, flipH: false, flipV: false }, coordinateSize: parentSize, children: independent };
  };
  r.known(tree); meta(tree, true);
  const props = r.one(tree, DSP, "grpSpPr");
  if (props) r.known(props);
  const rootGeometry = props ? r.transform(props) : undefined;
  if (props && r.one(props, r.a, "xfrm") && !rootGeometry) return empty("Cached root transform lacks local offset/extent.", part);
  const fields = props ? groupFields(props, rootGeometry?.box ?? extent) : { origin: { x: 0, y: 0 }, coordinateSize: extent };
  if (!fields) return empty("Cached root group has an incomplete coordinate transform.", part);
  const result: PresentationElementInput[] = [];
  for (const node of visual) result.push(await parse(node, fields.origin, fields.coordinateSize));
  await r.audit(root, paths.get(root)!, (path, reason) => { partial = true; note({ part, path }, reason); }, checkpoint);
  const content: GroupElementInput = { id: `${id}-cache-root`, type: "group", box: rootGeometry?.box ?? { x: 0, y: 0, ...extent },
    ...(rootGeometry ? { transform: rootGeometry.transform } : {}), coordinateSize: fields.coordinateSize, children: result };
  return { input: { ...base, type: "group", coordinateSize: extent, children: [content] }, status: partial ? "partial-cache" : "cache", cachePart: part };
}
