import type { Box } from "@pptkit/core";
import type { GridComposition, GridTrack } from "../types/composition.js";

interface Requirement { start: number; span: number; size: number }
const epsilon = (value: number): number => Number.EPSILON * 16 * Math.max(1, Math.abs(value));

/** Bounded, monotonic track solver. Returns undefined without mutating inputs on conflict. */
export function solveTracks(tracks: readonly GridTrack[], available: number, gap: number, requirements: readonly Requirement[]): number[] | undefined {
  const sizes = tracks.map((track) => track.type === "fixed" ? track.size : track.min ?? 0);
  const caps = tracks.map((track) => track.type === "fixed" ? track.size : track.max ?? Infinity);
  const budget = available - gap * (tracks.length - 1);
  if (!Number.isFinite(budget) || budget < 0) return undefined;
  // Interval demands are processed by their right edge. Grow rightmost capacity
  // first: it remains available to later overlapping intervals and avoids the
  // false overflow produced by equal growth of overlapping spans.
  const ordered = [...requirements].sort((a, b) => (a.start + a.span) - (b.start + b.span) || b.start - a.start || a.size - b.size);
  for (const requirement of ordered) {
    let remaining = requirement.size - gap * (requirement.span - 1);
    for (let i = requirement.start; i < requirement.start + requirement.span; i++) remaining -= sizes[i]!;
    for (let i = requirement.start + requirement.span - 1; i >= requirement.start && remaining > epsilon(requirement.size); i--) {
      const increment = Math.min(remaining, caps[i]! - sizes[i]!);
      sizes[i]! += increment; remaining -= increment;
    }
    if (remaining > epsilon(requirement.size)) return undefined;
  }
  let free = budget - sizes.reduce((sum, size) => sum + size, 0);
  if (!Number.isFinite(free) || free < -epsilon(available)) return undefined;
  while (free > epsilon(available)) {
    const eligible = tracks.flatMap((track, i) => track.type === "fraction" && sizes[i]! < caps[i]! ? [i] : []);
    if (!eligible.length) break;
    // Normalize first to avoid overflowing when valid finite weights have a huge sum.
    const weight = (i: number): number => { const track = tracks[i]!; return track.type === "fraction" ? track.weight ?? 1 : 1; };
    const largest = eligible.reduce((max, i) => Math.max(max, weight(i)), 0);
    const weights = eligible.map((i) => weight(i) / largest);
    const total = weights.reduce((sum, weight) => sum + weight, 0);
    let growth = 0;
    eligible.forEach((i, index) => {
      const increment = Math.min(free * (weights[index]! / total), caps[i]! - sizes[i]!);
      sizes[i]! += increment; growth += increment;
    });
    if (!(growth > 0)) break;
    free -= growth;
  }
  return sizes.every(Number.isFinite) ? sizes : undefined;
}

export function gridTargets(rule: GridComposition, ids: readonly string[], boxes: readonly Box[], area: Box): Box[] | undefined {
  const placements = new Map(rule.placements.map((placement) => [placement.elementId, placement]));
  const selected = ids.map((id) => placements.get(id)!);
  const columns = solveTracks(rule.columns, area.width, rule.columnGap ?? 0, selected.map((p, i) => ({ start: p.column, span: p.columnSpan ?? 1, size: boxes[i]!.width })));
  const rows = solveTracks(rule.rows, area.height, rule.rowGap ?? 0, selected.map((p, i) => ({ start: p.row, span: p.rowSpan ?? 1, size: boxes[i]!.height })));
  if (!columns || !rows) return undefined;
  const sum = (values: readonly number[], start: number, count: number): number => values.slice(start, start + count).reduce((a, b) => a + b, 0);
  const offset = (space: number, align: string | undefined): number => align === "end" ? space : align === "center" ? space / 2 : 0;
  return selected.map((p, i) => {
    const box = boxes[i]!;
    const width = sum(columns, p.column, p.columnSpan ?? 1) + (rule.columnGap ?? 0) * ((p.columnSpan ?? 1) - 1);
    const height = sum(rows, p.row, p.rowSpan ?? 1) + (rule.rowGap ?? 0) * ((p.rowSpan ?? 1) - 1);
    return { ...box,
      x: area.x + sum(columns, 0, p.column) + (rule.columnGap ?? 0) * p.column + offset(width - box.width, p.horizontal),
      y: area.y + sum(rows, 0, p.row) + (rule.rowGap ?? 0) * p.row + offset(height - box.height, p.vertical),
    };
  });
}

export function validateGrid(rule: GridComposition): void {
  for (const tracks of [rule.columns, rule.rows]) {
    if (!Array.isArray(tracks) || !tracks.length) throw new TypeError("Grid requires explicit nonempty tracks");
    for (const track of tracks) {
      if (!track || !["fixed", "content", "fraction"].includes(track.type)) throw new TypeError("Invalid grid track");
      if (track.type === "fixed") {
        if (!Number.isFinite(track.size) || track.size < 0) throw new RangeError("Fixed track size must be finite and nonnegative");
      } else {
        for (const value of [track.min, track.max]) if (value !== undefined && (!Number.isFinite(value) || value < 0)) throw new RangeError("Grid bounds must be finite and nonnegative");
        if ((track.min ?? 0) > (track.max ?? Infinity)) throw new RangeError("Grid minimum exceeds maximum");
        if (track.weight !== undefined && (track.type !== "fraction" || !Number.isFinite(track.weight) || track.weight <= 0)) throw new RangeError("Only fraction tracks accept positive finite weights");
      }
    }
  }
  for (const gap of [rule.columnGap, rule.rowGap]) if (gap !== undefined && (!Number.isFinite(gap) || gap < 0)) throw new RangeError("Grid gaps must be finite and nonnegative");
  if (!Array.isArray(rule.placements)) throw new TypeError("Grid requires placements");
  const ids = new Set<string>();
  for (const placement of rule.placements) {
    if (!placement || typeof placement.elementId !== "string" || ids.has(placement.elementId)) throw new TypeError("Grid placements require unique element IDs");
    ids.add(placement.elementId);
    for (const [start, span, length] of [[placement.column, placement.columnSpan ?? 1, rule.columns.length], [placement.row, placement.rowSpan ?? 1, rule.rows.length]]) {
      if (!Number.isSafeInteger(start) || start! < 0 || !Number.isSafeInteger(span) || span! < 1 || start! + span! > length!) throw new RangeError("Grid placement exceeds explicit tracks");
    }
    for (const align of [placement.horizontal, placement.vertical]) if (align !== undefined && !["start", "center", "end"].includes(align)) throw new TypeError("Invalid grid alignment");
  }
}
