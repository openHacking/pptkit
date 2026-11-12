import type { NormalizedTextBullet } from "../types/style.js";

/** Format an explicit normalized marker without maintaining list continuation state. */
export function formatListMarker(bullet: NormalizedTextBullet): string {
  if (!bullet || typeof bullet !== "object") throw new TypeError("A normalized bullet is required.");
  if (bullet.type === "none") return "";
  if (bullet.type === "bullet") {
    if (typeof bullet.character !== "string" || bullet.character.length === 0) throw new TypeError("A nonempty bullet character is required.");
    return `${bullet.character} `;
  }
  if (bullet.type !== "number" || !Number.isSafeInteger(bullet.startAt) || bullet.startAt < 1) throw new RangeError("Numbered-list startAt must be a positive safe integer.");
  if (!["arabicPeriod", "arabicParen", "alphaLowerPeriod", "alphaUpperPeriod"].includes(bullet.style)) throw new TypeError("Unsupported numbered-list style.");
  let value = String(bullet.startAt);
  if (bullet.style === "alphaLowerPeriod" || bullet.style === "alphaUpperPeriod") {
    value = "";
    let number = bullet.startAt;
    while (number > 0) { number--; value = String.fromCharCode(97 + number % 26) + value; number = Math.floor(number / 26); }
    if (bullet.style === "alphaUpperPeriod") value = value.toUpperCase();
  }
  return `${value}${bullet.style === "arabicParen" ? ")" : "."} `;
}
