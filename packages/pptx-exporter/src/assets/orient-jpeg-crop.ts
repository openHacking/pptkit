import type { ImageCrop } from "@pptkit/core";

/** Map a raw JPEG pixel-space crop to the same region after applying EXIF orientation. */
export function orientJpegCrop(crop: ImageCrop, orientation: 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8): ImageCrop {
  if (!crop || typeof crop !== "object") throw new TypeError("Crop must be an object.");
  const { left, top, right, bottom } = crop;
  if (![left, top, right, bottom].every(value => typeof value === "number" && Number.isFinite(value) && value >= 0 && value <= 1)
    || left + right >= 1 || top + bottom >= 1) throw new RangeError("Crop must leave a positive visible region.");
  if (!Number.isInteger(orientation) || orientation < 1 || orientation > 8) throw new RangeError("JPEG orientation must be an integer from 1 to 8.");
  switch (orientation) {
    case 1: return { left, top, right, bottom };
    case 2: return { left: right, top, right: left, bottom };
    case 3: return { left: right, top: bottom, right: left, bottom: top };
    case 4: return { left, top: bottom, right, bottom: top };
    case 5: return { left: top, top: left, right: bottom, bottom: right };
    case 6: return { left: bottom, top: left, right: top, bottom: right };
    case 7: return { left: bottom, top: right, right: top, bottom: left };
    case 8: return { left: top, top: right, right: bottom, bottom: left };
  }
}
