import type { Box, Point } from "@pptkit/core";
import type { LayoutResult } from "@pptkit/layout";
import { createSvgAffineMapping, createSvgViewportMapping, getSvgSlideSelectionFrames, getSvgSlideSelectionGeometry, hitTestSvgSelectionFrames, hitTestSvgSelectionOutlines } from "@pptkit/svg-renderer";

/** Client-space hit for an untransformed SVG content viewport. */
export function selectAt(layout: LayoutResult, slideIndex: number, viewport: Box, clientPoint: Point) {
  const mapping = createSvgViewportMapping(layout.size, viewport);
  const slidePoint = mapping.viewportToSlide(clientPoint);
  if (slidePoint.x < 0 || slidePoint.y < 0 || slidePoint.x > layout.size.width || slidePoint.y > layout.size.height) return undefined;
  const hit = hitTestSvgSelectionFrames(getSvgSlideSelectionFrames(layout, slideIndex), slidePoint);
  return hit && { elementId: hit.elementId, owner: hit.owner, ancestors: hit.ancestorIds };
}

/** Caller supplies the root SVG's current 2D screen matrix and client-space point. */
export function selectAtScreenMatrix(layout: LayoutResult, slideIndex: number, root: SVGSVGElement, clientPoint: Point) {
  const matrix = root.getScreenCTM();
  if (!matrix) return undefined;
  const point = createSvgAffineMapping(matrix).viewportToSlide(clientPoint);
  return hitTestSvgSelectionFrames(getSvgSlideSelectionFrames(layout, slideIndex), point);
}

/** A coarse visible-object policy; painted-pixel testing remains application-owned. */
export function selectNonzeroOpacity(layout: LayoutResult, slideIndex: number, point: Point, connectorTolerance: number) {
  const geometry = getSvgSlideSelectionGeometry(layout, slideIndex, { excludeZeroOpacity: true, excludeUnpaintedShapes: true });
  return hitTestSvgSelectionOutlines(geometry, point, connectorTolerance);
}
