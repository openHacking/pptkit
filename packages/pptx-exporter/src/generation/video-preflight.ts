import type { NormalizedElement, NormalizedPresentation } from "@pptkit/core";

/** Do not silently omit video or resolve any assets while video output is unavailable. */
export function rejectVideoOutput(document: NormalizedPresentation): void {
  const stack: NormalizedElement[] = [
    ...document.slides.flatMap(slide => slide.elements),
    ...document.layouts.flatMap(layout => layout.elements),
    ...(document.masters ?? []).flatMap(master => master.elements ?? []),
  ];
  while (stack.length) {
    const element = stack.pop()!;
    if (element.type === "video") throw new TypeError("PPTX video export is not supported; video cannot be silently omitted.");
    if (element.type === "group") for (const child of element.children) stack.push(child);
  }
}
