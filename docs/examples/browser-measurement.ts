import { createPresentation } from "@pptkit/core";
import { resolveLayout } from "@pptkit/layout";
import { createBrowserTextMeasurement } from "@pptkit/layout/dom";

// The application supplies licensed font bytes and their verified cmap coverage.
export async function measureInBrowser(fontBytes: ArrayBuffer, codePoints: ReadonlySet<number>) {
  const face = await new FontFace("Application Font", fontBytes).load();
  document.fonts.add(face);
  const adapter = createBrowserTextMeasurement({
    document,
    fonts: [{
      family: "Application Font",
      faces: [face],
      covers: ({ text, bold, italic }) => !bold && !italic && [...text].every(
        (character) => /[\n\r\t]/.test(character) || codePoints.has(character.codePointAt(0)!),
      ),
    }],
  });
  try {
    const presentation = createPresentation({ theme: { fonts: { body: "Application Font" } } });
    presentation.addSlide().addElement({
      type: "text", content: "Measured text", box: { x: 20, y: 20, width: 100, height: 30 },
    });
    return resolveLayout(presentation, {
      textMeasurement: adapter,
      diagnostics: { textOverflow: true, slideBounds: true },
    });
  } finally {
    adapter.dispose();
    document.fonts.delete(face);
  }
}
