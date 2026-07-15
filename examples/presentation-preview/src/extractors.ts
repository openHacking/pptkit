import { measureImageDimensions, type SourceParsers } from "@pptkit/presentation-workflow";
import pdfWorkerUrl from "pdfjs-dist/build/pdf.worker.min.mjs?url";

export const browserSourceParsers: SourceParsers = {
  async pdf(input) {
    const pdfjs = await import("pdfjs-dist/legacy/build/pdf.mjs");
    pdfjs.GlobalWorkerOptions.workerSrc = pdfWorkerUrl;
    const pdf = await pdfjs.getDocument({ data: input.bytes, useWorkerFetch: false, isEvalSupported: false, useSystemFonts: true }).promise;
    const pages: string[] = [];
    for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
      const page = await pdf.getPage(pageNumber);
      const text = await page.getTextContent();
      pages.push(text.items.map((item) => "str" in item ? item.str : "").join(" "));
    }
    return { content: pages.join("\n\n") };
  },
  async docx(input) {
    const mammoth = await import("mammoth/mammoth.browser");
    const result = await mammoth.extractRawText({ arrayBuffer: input.bytes.slice().buffer });
    return { content: result.value, warnings: result.messages.map((message) => message.message) };
  },
  async workbook(input) {
    const XLSX = await import("xlsx");
    const workbook = XLSX.read(input.bytes, { type: "array" });
    return {
      sheets: workbook.SheetNames.map((name) => ({
        name,
        rows: XLSX.utils.sheet_to_json(workbook.Sheets[name]!, { header: 1, raw: false }) as unknown[][],
      })),
    };
  },
  async image(input) {
    return measureImageDimensions(input);
  },
};
