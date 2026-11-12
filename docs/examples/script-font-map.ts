import { createPresentation, normalizePresentation, resolveTextFontSegments, type ScriptFontFamiliesInput } from "@pptkit/core";
import { generatePptx } from "@pptkit/pptx-exporter";

const scriptFontFamilies: ScriptFontFamiliesInput = {
  Hani: "Noto Sans CJK SC", Hant: "Noto Sans CJK TC",
  Arab: "Noto Sans Arabic", Hebr: "Noto Sans Hebrew",
};
const p = createPresentation();
p.addSlide().addElement({ type: "text", box: { x: 20, y: 20, width: 500, height: 90 }, content: [{ runs: [{
  text: "中文 العربية עברית", style: { fontFamily: "Noto Sans", language: "zh-TW", scriptFontFamilies },
}] }] });
const normalized = normalizePresentation(p), run = normalized.slides[0]!.elements[0]!;
if (run.type === "text") {
  const first = run.content[0]!.runs[0]!;
  console.log(resolveTextFontSegments(first.text, first.style, normalized.theme));
}
const output = await generatePptx(p);
// Consumers must inspect materialization warnings; font names do not install fonts.
console.log(output.warnings, output.bytes.length);
