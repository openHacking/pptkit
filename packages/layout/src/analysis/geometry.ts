import type { Box, NormalizedTransform, Point } from "@pptkit/core";

export type Matrix = readonly [number, number, number, number, number, number];
export const identity: Matrix = [1, 0, 0, 1, 0, 0];

export function multiply(a: Matrix, b: Matrix): Matrix {
  return [a[0] * b[0] + a[2] * b[1], a[1] * b[0] + a[3] * b[1],
    a[0] * b[2] + a[2] * b[3], a[1] * b[2] + a[3] * b[3],
    a[0] * b[4] + a[2] * b[5] + a[4], a[1] * b[4] + a[3] * b[5] + a[5]];
}

export function elementTransform(box: Box, transform: NormalizedTransform): Matrix {
  const angle = transform.rotation * Math.PI / 180;
  const cos = Math.cos(angle);
  const sin = Math.sin(angle);
  const x = box.x + box.width / 2;
  const y = box.y + box.height / 2;
  const fx = transform.flipH ? -1 : 1;
  const fy = transform.flipV ? -1 : 1;
  const a = cos * fx;
  const b = sin * fx;
  const c = -sin * fy;
  const d = cos * fy;
  return [a, b, c, d, x - a * x - c * y, y - b * x - d * y];
}

function point(matrix: Matrix, x: number, y: number): Point {
  return { x: matrix[0] * x + matrix[2] * y + matrix[4], y: matrix[1] * x + matrix[3] * y + matrix[5] };
}

export function transformedBounds(box: Box, matrix: Matrix): Box {
  const points = [point(matrix, box.x, box.y), point(matrix, box.x + box.width, box.y),
    point(matrix, box.x, box.y + box.height), point(matrix, box.x + box.width, box.y + box.height)];
  const x = Math.min(...points.map((p) => p.x));
  const y = Math.min(...points.map((p) => p.y));
  return { x, y, width: Math.max(...points.map((p) => p.x)) - x, height: Math.max(...points.map((p) => p.y)) - y };
}

export function overflows(bounds: Box, available: Box, tolerance: number): boolean {
  return bounds.x < available.x - tolerance || bounds.y < available.y - tolerance
    || bounds.x + bounds.width > available.x + available.width + tolerance
    || bounds.y + bounds.height > available.y + available.height + tolerance;
}
