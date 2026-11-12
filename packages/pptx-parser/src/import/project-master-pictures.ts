import { readParallelogramSlant } from "./parallelogram-slant.js";
import type { BackgroundPaintInput, ManagedPresentationDocument, PaintInput, PresentationElementInput, ShapeKind } from "@pptkit/core";
import { attribute, children, integer, invalidStructure } from "../presentation/structure-xml.js";
import type { structureContext } from "../presentation/structure-context.js";
import type { XmlElement } from "../xml/read-xml.js";
import { PptxImportError, type ImportFinding, type ImportSourceIdentity, type PptxSourceLocation } from "../types/import.js";
import { contentPaths, contentReader } from "./content-reader.js";
import { readDecorative } from "./decorative.js";
import type { contentImages } from "./content-images.js";
import type { MasterOwner } from "./project-structure.js";
import { contentText, type ContentTextInheritance } from "./content-text.js";
import type { ContentReader } from "./content-reader.js";
import { artworkPaintReader } from "./artwork-paint-reader.js";
import { tableStyleProperties } from "./table-style-properties.js";
import type { ThemeTextScope } from "./theme-text-resolver.js";

const shapePresets: Readonly<Record<string, ShapeKind>> = { rect: "rect", roundRect: "roundRect", ellipse: "ellipse", triangle: "triangle", diamond: "diamond", rightArrow: "arrow", chevron: "chevron", parallelogram: "parallelogram" };

/** Explicit partial projection of shared pictures, selected plain shapes and their groups. */
export async function projectMasterPictures(
  document: ManagedPresentationDocument, owners: readonly MasterOwner[], context: Awaited<ReturnType<typeof structureContext>>,
  image: ReturnType<typeof contentImages>, identities: ImportSourceIdentity[], checkpoint: () => Promise<void>,
  finding: (feature: string, operation: ImportFinding["operation"], tier: ImportFinding["tier"], source: PptxSourceLocation, reason: string, slideId?: string, elementId?: string, resolvedFrom?: readonly PptxSourceLocation[]) => void, includeShapes: boolean, includeText: boolean,
  textContext?: (owner: MasterOwner, r: ContentReader, paths: WeakMap<XmlElement, string>, source: PptxSourceLocation) => ContentTextInheritance,
  includeThemePaint = false, maxPaintDependencies = 1_000_000, includeBackground = false, includeBackgroundMatrix = false, includeBackgroundGradient = false, includeShapeGradient = false, includeShapePattern = false, includeShapeShadow = false,
): Promise<void> {
  let paintDependencies = 0;
  const slides = new Map(document.slides.map(slide => [slide.id, slide]));
  for (const owner of owners) {
    await checkpoint();
    const root = await context.root(owner.part, owner.kind === "master" ? "sldMaster" : owner.kind === "layout" ? "sldLayout" : "sld"), r = contentReader(owner.part, context.p, context.a);
    const paths = await contentPaths(root, checkpoint), inputs: PresentationElementInput[] = [];
    r.known(root); r.attr(root, "showMasterSp");
    const common = r.one(root, r.p, "cSld"), tree = common && r.one(common, r.p, "spTree");
    if (common) { r.known(common); r.attr(common, "name"); }
    const { paint, gradient, pattern, shadow: readShadow } = artworkPaintReader(r, checkpoint, (color, source) => textContext!(owner, r, paths, source).local.resolveColor!(color), (count, source) => {
      if (count > maxPaintDependencies - paintDependencies) throw new PptxImportError("import-limit-exceeded", "Shared artwork paint dependencies exceed maxTextStyleDependencies.", source);
      paintDependencies += count;
    });
    let background: BackgroundPaintInput | undefined, backgroundDependencies: readonly PptxSourceLocation[] | undefined;
    const bg = includeBackground && common ? r.one(common, r.p, "bg") : undefined;
    if (bg) {
      r.known(bg);
      const at = { part: owner.part, path: paths.get(bg)! }, properties = r.one(bg, r.p, "bgPr"), reference = r.one(bg, r.p, "bgRef");
      if (properties && reference) invalidStructure("Conflicting background properties and style reference.", owner.part);
      if (properties) { r.known(properties); const candidate = includeBackgroundGradient && r.one(properties, r.a, "gradFill") ? await gradient(properties, at, node => { const path = paths.get(node); return path ? { part: owner.part, path } : undefined; }, locations => { backgroundDependencies = locations; }) : await paint(properties, at); if (candidate?.type === "solid" || candidate?.type === "linearGradient") background = candidate; }
      if (reference && includeBackgroundMatrix) {
        const scope = textContext!(owner, r, paths, at).local.themeScope as ThemeTextScope;
        const resolver = tableStyleProperties(r.a, checkpoint, (source, reason) => finding("artwork.background.matrix-unmapped-property", "import", "unsupported", source, reason, owner.kind === "slide" ? owner.id : undefined));
        const selected = await resolver.matrixReference(reference, "fill", { part: owner.part, r, paths }, scope);
        const count = selected.resolvedFrom?.length ?? 0;
        if (count > maxPaintDependencies - paintDependencies) throw new PptxImportError("import-limit-exceeded", "Shared artwork paint dependencies exceed maxTextStyleDependencies.", at);
        paintDependencies += count;
        backgroundDependencies = selected.resolvedFrom;
        const candidate = selected.status === "mapped" ? selected.value as PaintInput : undefined;
        if (candidate?.type === "solid") background = candidate;
        else if (selected.status === "mapped") r.note(reference, "No-fill background matrix entries are not projected.");
      }
      finding("artwork.background", "import", background ? "fallback" : "unsupported", at, background ? "Supported background paint is materialized; effects, source-preserving edits and page-specific inherited theme changes remain separate." : "Background is outside the supported profile; Core inheritance is a fallback and may differ from the source.", owner.kind === "slide" ? owner.id : undefined, undefined, backgroundDependencies);
    }
    if (owner.kind === "slide") {
      if (background) slides.get(owner.id)!.setBackground(background);
      if (bg) await r.audit(bg, paths.get(bg)!, (path, reason) => finding("artwork.background.unmapped-property", "import", "unsupported", { part: owner.part, path }, reason, owner.id), checkpoint);
      continue;
    }
    // Validate identities even in omitted recognized drawings; opaque wrappers stay opaque.
    const seen = new Set<number>(), pending = tree ? [tree] : [];
    while (pending.length) {
      await checkpoint();
      const node = pending.pop()!;
      if (node.namespaceUri !== r.p) continue;
      const container = { spTree: "nvGrpSpPr", grpSp: "nvGrpSpPr", sp: "nvSpPr", pic: "nvPicPr", cxnSp: "nvCxnSpPr", graphicFrame: "nvGraphicFramePr" }[node.local];
      if (!container) continue;
      const nv = r.one(node, r.p, container), cnv = nv && r.one(nv, r.p, "cNvPr");
      if (cnv) {
        const id = integer(attribute(cnv, "id"), 0, 4294967295, "cNvPr/@id", owner.part);
        if (seen.has(id)) invalidStructure("Duplicate drawing identity in a master or layout part.", owner.part);
        seen.add(id);
      }
      if (node.local === "spTree" || node.local === "grpSp") for (const child of children(node)) pending.push(child);
    }
    const parse = async (node: XmlElement, origin: { x: number; y: number }, depth: number): Promise<PresentationElementInput | undefined> => {
        await checkpoint();
        if (node.namespaceUri !== r.p || (node.local !== "pic" && node.local !== "grpSp" && !(includeShapes && node.local === "sp"))) return undefined;
        const group = node.local === "grpSp", shape = node.local === "sp", feature = group ? "master-artwork.group" : shape ? "master-artwork.shape" : "master-artwork.picture";
        const sourceIndex = (await context.sourceElement(owner.part, node)).elementIndex;
        const nv = r.one(node, r.p, group ? "nvGrpSpPr" : shape ? "nvSpPr" : "nvPicPr"), cnv = nv && r.one(nv, r.p, "cNvPr"), nvPr = nv && r.one(nv, r.p, "nvPr");
        const at: PptxSourceLocation = { part: owner.part, path: paths.get(node)!, elementIndex: sourceIndex, ...(cnv && attribute(cnv, "id") !== undefined ? { sourceObjectId: attribute(cnv, "id")! } : {}) };
        const id = `${owner.id}-${group ? "group" : shape ? "shape" : "picture"}-${sourceIndex}`;
        const reject = (reason: string) => { finding(feature, "import", "unsupported", at, reason); return undefined; };
        if (group && depth >= 48) return reject("Group nesting exceeds the 48-level projection limit.");
        if (!nv || !cnv || attribute(cnv, "id") === undefined || (nvPr && r.one(nvPr, r.p, "ph"))) return reject("Missing drawing identity or placeholder: ordinary artwork cannot be inferred.");
        const properties = r.one(node, r.p, group ? "grpSpPr" : "spPr"), geometry = properties && r.transform(properties);
        if (!properties || !geometry || geometry.box.width <= 0 || geometry.box.height <= 0) return reject("An explicit positive frame is required.");
        geometry.box.x -= origin.x; geometry.box.y -= origin.y;
        r.known(node); r.known(nv); r.known(properties); r.attr(cnv, "id");
        const name = r.attr(cnv, "name") ?? "", description = r.attr(cnv, "descr"), hidden = r.bool(cnv, "hidden"), decorative = readDecorative(cnv, r);
        const objectPr = r.one(nv, r.p, group ? "cNvGrpSpPr" : shape ? "cNvSpPr" : "cNvPicPr"); if (objectPr) r.known(objectPr); if (nvPr) r.known(nvPr);
        const base = { id, name, hidden, ...geometry, ...(description !== undefined || decorative !== undefined ? { accessibility: { ...(description !== undefined ? { description } : {}), ...(decorative !== undefined ? { decorative: decorative.value } : {}) } } : {}) };
        let input: PresentationElementInput;
        if (group) {
          const xfrm = r.one(properties, r.a, "xfrm")!, off = r.one(xfrm, r.a, "chOff"), ext = r.one(xfrm, r.a, "chExt");
          if (!off || !ext) return reject("Group child offset and extent are required.");
          const childOrigin = { x: r.number(off, "x", -27273042329600, 27273042316900) / 12700, y: r.number(off, "y", -27273042329600, 27273042316900) / 12700 };
          const coordinateSize = { width: r.number(ext, "cx", 0, 27273042316900) / 12700, height: r.number(ext, "cy", 0, 27273042316900) / 12700 };
          if (!coordinateSize.width || !coordinateSize.height) return reject("Zero group child extent cannot form an invertible coordinate space.");
          const nested: PresentationElementInput[] = [];
          for (const child of children(node)) { const mapped = await parse(child, childOrigin, depth + 1); if (mapped) nested.push(mapped); }
          if (!nested.length) return reject("Group has no mapped artwork descendants; omitted content is reported separately.");
          input = { ...base, type: "group", coordinateSize, children: nested };
        } else if (shape) {
          const body = r.one(node, r.p, "txBody");
          if (body && !includeText) return reject("Shape text is not mapped by the shared artwork profile.");
          const preset = r.one(properties, r.a, "prstGeom"), custom = r.one(properties, r.a, "custGeom");
          const kind = preset && attribute(preset, "prst"), adjustments = preset && r.one(preset, r.a, "avLst");
          const slant = kind === "parallelogram" && preset ? readParallelogramSlant(preset, r.a) : undefined;
          if (!kind || !Object.hasOwn(shapePresets, kind) || custom || (slant ? !slant.supported : adjustments && children(adjustments).length > 0)) return reject("Shape geometry or adjustments are outside the supported preset profile.");
          const line = r.one(properties, r.a, "ln"); if (line && includeThemePaint) r.known(line);
          let shapeGradientDependencies: readonly PptxSourceLocation[] | undefined;
          const fill = includeShapeGradient && r.one(properties, r.a, "gradFill") ? await gradient(properties, at, node => { const path = paths.get(node); return path ? { part: owner.part, path } : undefined; }, locations => { shapeGradientDependencies = locations; }) : includeShapePattern && r.one(properties, r.a, "pattFill") ? await pattern(properties, at, node => { const path = paths.get(node); return path ? { part: owner.part, path } : undefined; }, locations => { shapeGradientDependencies = locations; }) : includeThemePaint ? await paint(properties, at) : r.paint(properties), linePaint = line && (includeThemePaint ? await paint(line, at) : r.paint(line));
          if (!fill || !line || !linePaint || (linePaint.type !== "none" && attribute(line, "w") === undefined)) return reject("Explicit supported shape fill and line are required; inherited paint is not inferred.");
          r.attr(preset!, "prst"); if (adjustments) r.known(adjustments);
          if (slant?.supported && slant.guide) { r.known(slant.guide); r.attr(slant.guide, "name"); r.attr(slant.guide, "fmla"); }
          if (shapeGradientDependencies) finding(fill.type === "pattern" ? "master-artwork.pattern-dependencies" : "master-artwork.gradient-dependencies", "import", "fallback", at, "Selected fill and resolved color dependencies; this is not a complete source reference graph.", undefined, id, shapeGradientDependencies);
          let shadowDependencies: readonly PptxSourceLocation[] | undefined;
          const shadow = includeShapeShadow ? await readShadow(properties, at, node => { const path = paths.get(node); return path ? { part: owner.part, path } : undefined; }, locations => { shadowDependencies = locations; }) : undefined;
          if (shadowDependencies) finding("master-artwork.shadow-dependencies", "import", "fallback", at, "Selected explicit shadow and resolved color dependencies; this is not a complete source reference graph.", undefined, id, shadowDependencies);
          const stroke = { ...r.stroke(properties)!, paint: linePaint };
          const text = body ? await contentText(body, r, checkpoint, () => undefined, true, textContext?.(owner, r, paths, at)) : undefined;
          if (body) finding("master-artwork.shape-text", "import", "fallback", at, `Direct text fields${textContext ? " and supported owner-scoped theme references" : ""} are projected without master/layout text-style inheritance, fields, actions or source-preserving edit bindings; page-specific theme overrides on shared artwork are not modeled. Unhandled nodes are reported separately.`, undefined, id);
          input = { ...base, type: "shape", shape: shapePresets[kind]!, ...(slant?.supported && slant.value !== undefined ? { slant: slant.value } : {}), style: { fill, stroke, ...(shadow ? { shadow } : {}) }, ...(text ? { text } : {}) };
        } else {
          const mapped = await image(node, properties, r, at);
          if ("reason" in mapped) return reject(mapped.reason);
          input = { ...base, ...mapped.fields };
        }
        identities.push({ kind: "element", id, source: at });
        if (decorative) finding("master-artwork.decorative", "import", "editable", { ...at, path: paths.get(decorative.node)! }, "Explicit decorative metadata is projected; source edit-save and application behavior require separate verification.", undefined, id);
        finding(feature, "import", "fallback", at, "Shared artwork fields are projected; unknown properties and surrounding artwork are separately reported. Source-preserving edits are not supported.", undefined, id);
        return input;
    };
    if (tree) {
      r.known(tree);
      for (const node of children(tree)) {
        const mapped = await parse(node, { x: 0, y: 0 }, 0);
        if (mapped) inputs.push(mapped);
      }
    }
    if (owner.kind === "master") document.replaceSlideMaster(owner.id, { name: owner.name, elements: inputs, ...(background ? { background } : {}) });
    else document.replaceSlideLayout(owner.id, { name: owner.name, masterId: owner.masterId!, elements: inputs, ...(background ? { background } : {}), ...(owner.showMasterShapes !== undefined ? { showMasterShapes: owner.showMasterShapes } : {}) });
    await r.audit(root, paths.get(root)!, (path, reason) => finding("master-artwork.unmapped-property", "import", "unsupported", { part: owner.part, path }, reason), checkpoint);
  }
}
