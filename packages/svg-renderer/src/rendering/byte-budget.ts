import { SvgRenderLimitError } from '../types/public.js';

/** Counts UTF-8 without allocating a byte buffer; unpaired surrogates use U+FFFD's three bytes. */
export function chargeSvgBytes(value: string, used: number, limit: number, slideIndex: number): number {
  for (const char of value) {
    const point = char.codePointAt(0)!;
    const bytes = point <= 0x7f ? 1 : point <= 0x7ff ? 2 : point <= 0xffff ? 3 : 4;
    if (used > limit - bytes) throw new SvgRenderLimitError(limit, slideIndex);
    used += bytes;
  }
  return used;
}
