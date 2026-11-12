import type { Box, Point, Size } from "@pptkit/core";

/** Mapping for an untransformed SVG viewport using the renderer's default xMidYMid meet. */
export interface SvgViewportMapping {
  readonly scale: number;
  readonly contentBox: Readonly<Box>;
  slideToViewport(point: Point): Readonly<Point>;
  viewportToSlide(point: Point): Readonly<Point>;
}

/** SVG/DOMMatrix-compatible 2D affine coefficients. */
export interface SvgAffineMatrix {
  readonly a: number; readonly b: number; readonly c: number;
  readonly d: number; readonly e: number; readonly f: number;
  readonly is2D?: boolean;
}
export interface SvgAffineMapping {
  readonly matrix: Readonly<SvgAffineMatrix>;
  slideToViewport(point: Point): Readonly<Point>;
  viewportToSlide(point: Point): Readonly<Point>;
}

/** Map slide coordinates using a caller-supplied, invertible SVG/DOM 2D matrix. */
export function createSvgAffineMapping(input: SvgAffineMatrix): SvgAffineMapping {
  const { a, b, c, d, e, f, is2D } = input;
  if (is2D !== undefined && is2D !== true) throw new TypeError("SVG affine mapping requires a 2D matrix.");
  if (![a, b, c, d, e, f].every(Number.isFinite)) throw new TypeError("SVG affine matrix coefficients must be finite.");
  const scale = Math.max(Math.abs(a), Math.abs(b), Math.abs(c), Math.abs(d));
  if (scale === 0) throw new RangeError("SVG affine matrix must be invertible.");
  const na = a / scale, nb = b / scale, nc = c / scale, nd = d / scale;
  const determinant = na * nd - nb * nc;
  if (!Number.isFinite(determinant) || determinant === 0) throw new RangeError("SVG affine matrix must be numerically invertible.");
  const ia = nd / determinant / scale, ib = -nb / determinant / scale, ic = -nc / determinant / scale, id = na / determinant / scale;
  if (![ia, ib, ic, id].every(Number.isFinite)) throw new RangeError("SVG inverse affine matrix cannot be represented with finite coefficients.");
  const matrix = Object.freeze({ a, b, c, d, e, f, is2D: true });
  const result = (x: number, y: number): Readonly<Point> => {
    if (!Number.isFinite(x) || !Number.isFinite(y)) throw new RangeError("Mapped SVG coordinates are not finite.");
    return Object.freeze({ x, y });
  };
  const check = (point: Point) => { if (!Number.isFinite(point.x) || !Number.isFinite(point.y)) throw new TypeError("SVG mapping point must be finite."); };
  return Object.freeze({
    matrix,
    slideToViewport(point: Point) { check(point); return result(a * point.x + c * point.y + e, b * point.x + d * point.y + f); },
    viewportToSlide(point: Point) { check(point); const x = point.x - e, y = point.y - f; return result(ia * x + ic * y, ib * x + id * y); },
  });
}

/** Construct a detached mapping in the coordinate system of the supplied viewport box. */
export function createSvgViewportMapping(slideSize: Size, viewport: Box): SvgViewportMapping {
  if (![slideSize.width, slideSize.height, viewport.x, viewport.y, viewport.width, viewport.height].every(Number.isFinite) || slideSize.width <= 0 || slideSize.height <= 0 || viewport.width <= 0 || viewport.height <= 0) throw new TypeError("SVG viewport and slide dimensions must be finite and positive.");
  const scale = Math.min(viewport.width / slideSize.width, viewport.height / slideSize.height);
  const width = slideSize.width * scale, height = slideSize.height * scale;
  const x = viewport.x + (viewport.width - width) / 2, y = viewport.y + (viewport.height - height) / 2;
  if (![scale, width, height, x, y, x + width, y + height].every(Number.isFinite) || scale <= 0 || width <= 0 || height <= 0) throw new RangeError("SVG viewport mapping cannot be represented with finite dimensions.");
  const contentBox = Object.freeze({ x, y, width, height });
  const result = (x: number, y: number): Readonly<Point> => {
    if (!Number.isFinite(x) || !Number.isFinite(y)) throw new RangeError("Mapped SVG coordinates are not finite.");
    return Object.freeze({ x, y });
  };
  const check = (point: Point) => { if (!Number.isFinite(point.x) || !Number.isFinite(point.y)) throw new TypeError("SVG mapping point must be finite."); };
  return Object.freeze({
    scale, contentBox,
    slideToViewport(point: Point) { check(point); return result(x + point.x * scale, y + point.y * scale); },
    viewportToSlide(point: Point) { check(point); return result((point.x - x) / scale, (point.y - y) / scale); },
  });
}
