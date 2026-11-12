import { CliError } from "./arguments.js";

/** Wrap already-rendered SVGs in page order, without injecting authored text. */
export function previewHtml(slides: readonly { svg: string; index: number; hidden: boolean }[], width: number, height: number, limit: number): string {
  const parts: string[] = []; let bytes = 0;
  const append = (value: string): void => {
    bytes += Buffer.byteLength(value);
    if (bytes > limit) throw new CliError("output-limit-exceeded", "HTML preview exceeds max-output-bytes.", 4);
    parts.push(value);
  };
  append('<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>PPTKit preview</title><style>body{margin:0;background:#eee;font-family:system-ui,sans-serif}main{max-width:1200px;margin:auto;padding:1rem}section{margin:0 0 1.5rem}iframe{display:block;width:100%;border:0;background:white}h2{font-size:1rem;font-weight:500}</style></head><body><main>');
  for (const slide of slides) {
    append(`<section data-slide="${slide.index + 1}" data-hidden="${slide.hidden}"><h2>Slide ${slide.index + 1}${slide.hidden ? " (hidden)" : ""}</h2>`);
    const document = `<!doctype html><html><head><meta charset="utf-8"><style>html,body{margin:0}body>svg{display:block;width:100%;height:auto}</style></head><body>${slide.svg}</body></html>`;
    const escaped = document.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
    append(`<iframe sandbox="" title="Slide ${slide.index + 1}" style="aspect-ratio:${width}/${height}" srcdoc="${escaped}"></iframe>`); append("</section>");
  }
  append("</main></body></html>");
  return parts.join("");
}
