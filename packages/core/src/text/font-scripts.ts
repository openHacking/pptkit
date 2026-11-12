/** Supported ISO 15924 font-selection keys; no locale maximization or host fonts. */
export const TEXT_FONT_SCRIPT_CODES = Object.freeze([
  "Zyyy", "Latn", "Grek", "Cyrl", "Hani", "Hans", "Hant", "Hira", "Kana", "Jpan", "Hang", "Kore", "Bopo", "Yiii",
  "Arab", "Hebr", "Syrc", "Thaa", "Nkoo", "Samr", "Mand", "Deva", "Beng", "Guru", "Gujr", "Orya", "Taml", "Telu", "Knda", "Mlym", "Sinh", "Thai", "Laoo", "Tibt", "Mymr", "Khmr", "Mong",
] as const);
export type TextFontScriptCode = typeof TEXT_FONT_SCRIPT_CODES[number];
export type TextFontScript = "default" | "eastAsian" | "complexScript";

const eastern = /[\p{Script_Extensions=Han}\p{Script_Extensions=Hiragana}\p{Script_Extensions=Katakana}\p{Script_Extensions=Hangul}\p{Script_Extensions=Bopomofo}\p{Script_Extensions=Yi}]/u;
const complex = /[\p{Script=Arabic}\p{Script=Hebrew}\p{Script=Syriac}\p{Script=Thaana}\p{Script=Nko}\p{Script=Samaritan}\p{Script=Mandaic}\p{Script=Devanagari}\p{Script=Bengali}\p{Script=Gurmukhi}\p{Script=Gujarati}\p{Script=Oriya}\p{Script=Tamil}\p{Script=Telugu}\p{Script=Kannada}\p{Script=Malayalam}\p{Script=Sinhala}\p{Script=Thai}\p{Script=Lao}\p{Script=Tibetan}\p{Script=Myanmar}\p{Script=Khmer}\p{Script=Mongolian}]/u;
const codes: readonly [TextFontScriptCode, RegExp][] = [
  ["Latn", /\p{Script=Latin}/u], ["Grek", /\p{Script=Greek}/u], ["Cyrl", /\p{Script=Cyrillic}/u],
  ["Hani", /\p{Script_Extensions=Han}/u], ["Hira", /\p{Script_Extensions=Hiragana}/u], ["Kana", /\p{Script_Extensions=Katakana}/u], ["Hang", /\p{Script_Extensions=Hangul}/u], ["Bopo", /\p{Script_Extensions=Bopomofo}/u], ["Yiii", /\p{Script_Extensions=Yi}/u],
  ["Arab", /\p{Script=Arabic}/u], ["Hebr", /\p{Script=Hebrew}/u], ["Syrc", /\p{Script=Syriac}/u], ["Thaa", /\p{Script=Thaana}/u], ["Nkoo", /\p{Script=Nko}/u], ["Samr", /\p{Script=Samaritan}/u], ["Mand", /\p{Script=Mandaic}/u],
  ["Deva", /\p{Script=Devanagari}/u], ["Beng", /\p{Script=Bengali}/u], ["Guru", /\p{Script=Gurmukhi}/u], ["Gujr", /\p{Script=Gujarati}/u], ["Orya", /\p{Script=Oriya}/u], ["Taml", /\p{Script=Tamil}/u], ["Telu", /\p{Script=Telugu}/u], ["Knda", /\p{Script=Kannada}/u], ["Mlym", /\p{Script=Malayalam}/u], ["Sinh", /\p{Script=Sinhala}/u],
  ["Thai", /\p{Script=Thai}/u], ["Laoo", /\p{Script=Lao}/u], ["Tibt", /\p{Script=Tibetan}/u], ["Mymr", /\p{Script=Myanmar}/u], ["Khmr", /\p{Script=Khmer}/u], ["Mong", /\p{Script=Mongolian}/u],
];

export function fontScriptBucket(character: string): TextFontScript {
  return eastern.test(character) ? "eastAsian" : complex.test(character) ? "complexScript" : "default";
}
export function fontScriptCode(character: string): TextFontScriptCode {
  return codes.find(([, pattern]) => pattern.test(character))?.[0] ?? "Zyyy";
}

/** Read only explicit locale components, never CLDR likely-subtag defaults. */
export function hanFontScripts(language: string | undefined): readonly TextFontScriptCode[] {
  if (!language) return ["Hani"];
  let locale: Intl.Locale;
  try { locale = new Intl.Locale(language); } catch { return ["Hani"]; }
  if (locale.script === "Hans" || locale.script === "Hant") return [locale.script, "Hani"];
  if (locale.script === "Jpan") return ["Jpan", "Hani"];
  if (locale.script === "Kore" || locale.script === "Hang") return ["Kore", "Hang", "Hani"];
  if (locale.script && locale.script !== "Hani") return ["Hani"];
  if (locale.language === "ja") return ["Jpan", "Hani"];
  if (locale.language === "ko") return ["Kore", "Hang", "Hani"];
  if (locale.language === "zh" || locale.language === "yue") {
    if (["CN", "SG"].includes(locale.region ?? "")) return ["Hans", "Hani"];
    if (["TW", "HK", "MO"].includes(locale.region ?? "")) return ["Hant", "Hani"];
  }
  return ["Hani"];
}

export function fontScriptCandidates(code: TextFontScriptCode, han: readonly TextFontScriptCode[]): readonly TextFontScriptCode[] {
  if (code === "Hani") return han;
  if (code === "Hira" || code === "Kana") return [code, "Jpan"];
  if (code === "Hang") return ["Hang", "Kore"];
  if (code === "Bopo" && han[0] === "Hant") return ["Bopo", "Hant"];
  return [code];
}
