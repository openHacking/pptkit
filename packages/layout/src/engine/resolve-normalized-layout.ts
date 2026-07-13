import type {
  Box,
  ConnectorAnchor,
  ConnectorEndpointInput,
  NormalizedAsset,
  NormalizedElement,
  NormalizedPresentation,
  Point,
} from "@pptkit/core";
import type { LayoutElement, LayoutGroupElement, LayoutResult } from "../types/layout.js";

function anchorPoint(box: Box, anchor: ConnectorAnchor): Point {
  if (anchor === "top") return { x: box.x + box.width / 2, y: box.y };
  if (anchor === "right") return { x: box.x + box.width, y: box.y + box.height / 2 };
  if (anchor === "bottom") return { x: box.x + box.width / 2, y: box.y + box.height };
  if (anchor === "left") return { x: box.x, y: box.y + box.height / 2 };
  return { x: box.x + box.width / 2, y: box.y + box.height / 2 };
}

function resolveEndpoint(endpoint: ConnectorEndpointInput, boxes: ReadonlyMap<string, Box>): Point {
  if (!("elementId" in endpoint)) return { ...endpoint };
  const box = boxes.get(endpoint.elementId);
  if (box === undefined) return { x: 0, y: 0 };
  return anchorPoint(box, endpoint.anchor ?? "center");
}

function connectorBox(points: Point[]): Box {
  const xs = points.map((point) => point.x);
  const ys = points.map((point) => point.y);
  const x = Math.min(...xs);
  const y = Math.min(...ys);
  return { x, y, width: Math.max(...xs) - x, height: Math.max(...ys) - y };
}

function resolveImage(element: Extract<NormalizedElement, { type: "image" }>, assets: ReadonlyMap<string, NormalizedAsset>): LayoutElement {
  const clone = structuredClone(element);
  const asset = assets.get(element.assetId);
  if (asset?.width === undefined || asset.height === undefined || element.box.width === 0 || element.box.height === 0) return clone;
  const sourceAspect = asset.width / asset.height;
  const targetAspect = element.box.width / element.box.height;
  if (element.fit === "contain") {
    if (sourceAspect > targetAspect) {
      const height = element.box.width / sourceAspect;
      clone.box = { x: element.box.x, y: element.box.y + (element.box.height - height) / 2, width: element.box.width, height };
    } else {
      const width = element.box.height * sourceAspect;
      clone.box = { x: element.box.x + (element.box.width - width) / 2, y: element.box.y, width, height: element.box.height };
    }
    clone.fit = "stretch";
  } else if (element.fit === "cover") {
    if (sourceAspect > targetAspect) {
      const visible = targetAspect / sourceAspect;
      clone.crop = { left: (1 - visible) / 2, right: (1 - visible) / 2, top: 0, bottom: 0 };
    } else {
      const visible = sourceAspect / targetAspect;
      clone.crop = { left: 0, right: 0, top: (1 - visible) / 2, bottom: (1 - visible) / 2 };
    }
    clone.fit = "crop";
  }
  return clone;
}

function collectBoxes(elements: readonly NormalizedElement[], destination: Map<string, Box>): void {
  for (const element of elements) destination.set(element.id, element.box);
}

function resolveElements(elements: readonly NormalizedElement[], assets: ReadonlyMap<string, NormalizedAsset>, additional: readonly NormalizedElement[] = []): LayoutElement[] {
  const boxes = new Map<string, Box>();
  collectBoxes(additional, boxes);
  collectBoxes(elements, boxes);
  return elements.map((element): LayoutElement => {
    if (element.type === "connector") {
      const start = resolveEndpoint(element.start, boxes);
      const end = resolveEndpoint(element.end, boxes);
      const route = element.route.map((point) => ({ ...point }));
      return { ...structuredClone(element), start, end, route, box: connectorBox([start, ...route, end]) };
    }
    if (element.type === "group") {
      const group: LayoutGroupElement = {
        ...structuredClone(element),
        children: resolveElements(element.children, assets),
      };
      return group;
    }
    if (element.type === "image") return resolveImage(element, assets);
    return structuredClone(element);
  });
}

export function resolveNormalizedLayout(normalized: NormalizedPresentation): LayoutResult {
  const assets = new Map(normalized.assets.map((asset) => [asset.id, asset]));
  const normalizedLayouts = new Map(normalized.layouts.map((layout) => [layout.id, layout]));
  const layouts = normalized.layouts.map((layout) => ({
    id: layout.id,
    name: layout.name,
    background: structuredClone(layout.background),
    elements: resolveElements(layout.elements, assets),
    placeholders: structuredClone(layout.placeholders),
  }));
  return {
    size: structuredClone(normalized.size),
    metadata: structuredClone(normalized.metadata),
    theme: structuredClone(normalized.theme),
    assets: structuredClone(normalized.assets),
    layouts,
    slides: normalized.slides.map((slide) => ({
      id: slide.id,
      layoutId: slide.layoutId,
      background: structuredClone(slide.background),
      backgroundSource: slide.backgroundSource,
      elements: resolveElements(slide.elements, assets, normalizedLayouts.get(slide.layoutId)?.elements ?? []),
      notes: structuredClone(slide.notes),
      hidden: slide.hidden,
      ...(slide.section !== undefined ? { section: slide.section } : {}),
      tags: [...slide.tags],
      customData: structuredClone(slide.customData),
    })),
    slideCount: normalized.slides.length,
    status: "resolved",
  };
}
