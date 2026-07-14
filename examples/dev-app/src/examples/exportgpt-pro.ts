import type { ExampleDefinition, ExampleElementSpec, ExampleInputData } from "../example-types.js";

const W = 960;
const H = 540;
const NAVY = "#093B3B";
const DEEP = "#062F30";
const PAPER = "#F7F3E9";
const INK = "#102C2C";
const MUTED = "#B7CCC5";
const ORANGE = "#F2A23A";
const CORAL = "#D96A52";
const PALE = "#E7E8D8";
const FONT = "Aptos";

type E = ExampleElementSpec;

function text(textValue: string, x: number, y: number, width: number, height: number, style: NonNullable<Extract<E, { type: "text" }>["style"]> = {}): E {
  return {
    type: "text",
    text: textValue,
    box: { x, y, width, height },
    style: { fontFamily: FONT, autoFit: { mode: "shrink", fontScale: 0.9 }, ...style },
  };
}

function rect(x: number, y: number, width: number, height: number, fill: string, stroke?: string, strokeWidth = 1): E {
  return { type: "shape", shape: "rect", box: { x, y, width, height }, style: { fill, ...(stroke ? { stroke, strokeWidth } : {}) } };
}

function roundRect(x: number, y: number, width: number, height: number, fill: string, stroke?: string): E {
  return { type: "shape", shape: "roundRect", box: { x, y, width, height }, style: { fill, ...(stroke ? { stroke, strokeWidth: 1 } : {}) } };
}

function line(x: number, y: number, width: number, color = MUTED, strokeWidth = 1): E {
  return { type: "shape", shape: "line", box: { x, y, width, height: 0 }, style: { stroke: color, strokeWidth } };
}

function base(slideNo: string, section: string, background = NAVY): E[] {
  return [
    rect(0, 0, 20, H, ORANGE),
    rect(W - 20, 0, 20, H, CORAL),
    text(section.toUpperCase(), 56, 18, 330, 20, { fontSize: 10, fontWeight: "bold", color: ORANGE }),
    text(slideNo, 855, 18, 50, 20, { fontSize: 11, fontWeight: "bold", color: MUTED, align: "right" }),
    line(56, 42, 849, "#2B5B58", 0.75),
    text("ExportGPT  |  Export ChatGPT Conversations", 56, 510, 430, 16, { fontSize: 9, color: MUTED }),
    text(slideNo, 855, 510, 50, 16, { fontSize: 9, color: MUTED, align: "right" }),
  ];
}

function title(value: string, x = 56, y = 78, width = 600, height = 68, color = PAPER): E {
  return text(value, x, y, width, height, { fontSize: 34, fontWeight: "bold", color, lineSpacing: 0.95 });
}

function body(value: string, x: number, y: number, width: number, height: number, color = MUTED): E {
  return text(value, x, y, width, height, { fontSize: 17, color, lineSpacing: 1.15 });
}

function label(value: string, x: number, y: number, width: number, color = ORANGE): E {
  return text(value, x, y, width, 20, { fontSize: 11, fontWeight: "bold", color });
}

function metric(value: string, caption: string, x: number, y: number, width: number): E[] {
  return [
    text(value, x, y, width, 45, { fontSize: 31, fontWeight: "bold", color: ORANGE }),
    text(caption, x, y + 48, width, 36, { fontSize: 14, color: MUTED, lineSpacing: 1.1 }),
  ];
}

const input: ExampleInputData = {
  title: "ExportGPT — Export ChatGPT Conversations",
  summary: "A ten-slide, fully editable product introduction inspired by the supplied reference deck's visual language.",
  size: { width: W, height: H, unit: "pt" },
  slides: [
    {
      id: "exportgpt-cover",
      title: "Cover",
      background: DEEP,
      elements: [
        rect(0, 0, 20, H, ORANGE),
        rect(W - 160, 0, 80, H, ORANGE),
        rect(W - 80, 0, 80, H, CORAL),
        roundRect(56, 78, 150, 28, ORANGE),
        text("CHROME  ·  EDGE", 72, 84, 120, 16, { fontSize: 10, fontWeight: "bold", color: INK, align: "center" }),
        text("ExportGPT", 56, 145, 490, 58, { fontSize: 48, fontWeight: "bold", color: PAPER }),
        text("Turn ChatGPT conversations into\nmaterials you can actually deliver", 56, 220, 555, 116, { fontSize: 34, fontWeight: "bold", color: PAPER, lineSpacing: 0.95 }),
        line(56, 365, 145, ORANGE, 3),
        body("One-click export · structure preserved · multi-format output\nBring conversations into reports, docs, and knowledge bases.", 56, 388, 390, 72, PAPER),
        text("ExportGPT  |  Product introduction", 56, 510, 360, 16, { fontSize: 9, color: MUTED }),
        text("01", 855, 510, 50, 16, { fontSize: 9, color: MUTED, align: "right" }),
      ],
    },
    {
      id: "exportgpt-judgment",
      title: "Core judgment",
      background: NAVY,
      elements: [
        ...base("02", "Core judgment"),
        title("AI conversations should not stop at the chat window"),
        body("When output needs to be saved, reused, shared, or delivered, the conversation needs a real document-style exit.", 56, 160, 520, 58),
        roundRect(56, 260, 250, 160, "#124B49"),
        label("01  Limited formats", 78, 282, 180),
        text("Official export leans\nheavily toward JSON", 78, 315, 190, 60, { fontSize: 23, fontWeight: "bold", color: PAPER, lineSpacing: 0.95 }),
        roundRect(335, 260, 250, 160, "#124B49"),
        label("02  Manual work", 357, 282, 180),
        text("Screenshots, copy-paste,\nreformatting", 357, 315, 190, 60, { fontSize: 23, fontWeight: "bold", color: PAPER, lineSpacing: 0.95 }),
        roundRect(614, 260, 250, 160, "#124B49"),
        label("03  Fragmented tools", 636, 282, 180),
        text("PDF, Word,\nHTML, each with its own flow", 636, 315, 190, 60, { fontSize: 23, fontWeight: "bold", color: PAPER, lineSpacing: 0.95 }),
        text("The problem is not a lack of content, but a lack of usable files.", 56, 454, 700, 28, { fontSize: 18, fontWeight: "bold", color: ORANGE }),
      ],
    },
    {
      id: "exportgpt-flow",
      title: "Product mechanism",
      background: NAVY,
      elements: [
        ...base("03", "Product mechanism"),
        title("One-click export that drops content straight into the workflow"),
        body("Stay on the ChatGPT page, pick the parts you need, and generate deliverables immediately after editing.", 56, 158, 560, 48),
        line(168, 302, 650, "#7AA29A", 2),
        roundRect(84, 250, 170, 100, ORANGE),
        roundRect(300, 250, 170, 100, CORAL),
        roundRect(516, 250, 170, 100, "#D6D7C2"),
        roundRect(732, 250, 170, 100, ORANGE),
        text("ChatGPT\nconversation", 110, 273, 118, 55, { fontSize: 22, fontWeight: "bold", color: INK, align: "center" }),
        text("Select\ncontent pieces", 326, 273, 118, 55, { fontSize: 22, fontWeight: "bold", color: PAPER, align: "center" }),
        text("Edit &\npreview", 542, 273, 118, 55, { fontSize: 22, fontWeight: "bold", color: INK, align: "center" }),
        text("Export\nwork files", 758, 273, 118, 55, { fontSize: 22, fontWeight: "bold", color: INK, align: "center" }),
        text("1", 150, 370, 40, 30, { fontSize: 20, fontWeight: "bold", color: ORANGE, align: "center" }),
        text("2", 366, 370, 40, 30, { fontSize: 20, fontWeight: "bold", color: CORAL, align: "center" }),
        text("3", 582, 370, 40, 30, { fontSize: 20, fontWeight: "bold", color: PALE, align: "center" }),
        text("4", 798, 370, 40, 30, { fontSize: 20, fontWeight: "bold", color: ORANGE, align: "center" }),
        text("Every step serves one purpose: keep good answers from needing one more manual move.", 56, 454, 700, 28, { fontSize: 18, fontWeight: "bold", color: PAPER }),
      ],
    },
    {
      id: "exportgpt-formats",
      title: "Format coverage",
      background: NAVY,
      elements: [
        ...base("04", "Format coverage"),
        title("Nine file formats, covering everything from saving to delivery"),
        body("The same conversation can find the right exit for very different work scenarios.", 56, 158, 520, 40),
        rect(56, 245, 245, 176, "#124B49"),
        rect(337, 245, 245, 176, "#124B49"),
        rect(618, 245, 245, 176, "#124B49"),
        label("Lightweight saving", 78, 268, 180),
        text("Text\nMarkdown\nJSON", 78, 302, 160, 92, { fontSize: 22, fontWeight: "bold", color: PAPER, lineSpacing: 1.05 }),
        label("Formal delivery", 359, 268, 180),
        text("PDF\nWord\nHTML", 359, 302, 160, 92, { fontSize: 22, fontWeight: "bold", color: PAPER, lineSpacing: 1.05 }),
        label("Data reuse", 640, 268, 180),
        text("Excel\nCSV\nScreenshot", 640, 302, 180, 92, { fontSize: 22, fontWeight: "bold", color: PAPER, lineSpacing: 1.05 }),
        text("From personal notes to team knowledge bases, you do not need a different tool just to change formats.", 56, 456, 690, 28, { fontSize: 18, fontWeight: "bold", color: ORANGE }),
      ],
    },
    {
      id: "exportgpt-features",
      title: "Core capabilities",
      background: NAVY,
      elements: [
        ...base("05", "Core capabilities"),
        title("Turn export into a truly controllable action"),
        body("ExportGPT does more than download files; it keeps content selection, structure preservation, and pre-delivery checks in one place.", 56, 158, 690, 48),
        roundRect(56, 250, 390, 62, "#124B49"),
        roundRect(56, 328, 390, 62, "#124B49"),
        roundRect(490, 250, 390, 62, "#124B49"),
        roundRect(490, 328, 390, 62, "#124B49"),
        label("01", 78, 270, 40), text("One-click export", 126, 267, 120, 24, { fontSize: 20, fontWeight: "bold", color: PAPER }),
        label("02", 78, 348, 40), text("Selective export", 126, 345, 170, 24, { fontSize: 20, fontWeight: "bold", color: PAPER }),
        label("03", 512, 270, 40), text("Edit & preview", 560, 267, 170, 24, { fontSize: 20, fontWeight: "bold", color: PAPER }),
        label("04", 512, 348, 40), text("Image and share-link export", 560, 345, 220, 24, { fontSize: 20, fontWeight: "bold", color: PAPER }),
        text("Keep code blocks, tables, paragraphs, lists, and images intact so the exported file stays usable.", 56, 456, 760, 28, { fontSize: 18, fontWeight: "bold", color: ORANGE }),
      ],
    },
    {
      id: "exportgpt-scenarios",
      title: "Work scenarios",
      background: NAVY,
      elements: [
        ...base("06", "Work scenarios"),
        title("One conversation can become many kinds of work output"),
        body("ExportGPT is not about archiving; it is about carrying AI output into the next step.", 56, 158, 590, 42),
        line(56, 245, 800, "#2B5B58", 1),
        label("Reports", 56, 270, 120),
        text("Pull conclusions out of the research process and package them as shareable PDF or Word files.", 185, 266, 610, 30, { fontSize: 18, color: PAPER }),
        line(56, 328, 800, "#2B5B58", 1),
        label("Meeting notes", 56, 353, 120),
        text("Keep highlights, action items, and context so the team can review and follow up easily.", 185, 349, 610, 30, { fontSize: 18, color: PAPER }),
        line(56, 411, 800, "#2B5B58", 1),
        label("Technical docs", 56, 436, 120),
        text("Code blocks, lists, and structured explanations survive intact and go straight into the knowledge base.", 185, 432, 610, 30, { fontSize: 18, color: PAPER }),
        text("Turn one-off answers into long-lived, searchable, reusable organizational assets.", 56, 486, 760, 24, { fontSize: 17, fontWeight: "bold", color: ORANGE }),
      ],
    },
    {
      id: "exportgpt-selection",
      title: "Selection and preview",
      background: NAVY,
      elements: [
        ...base("07", "Select and preview"),
        title("No need to move the whole conversation, export only what you actually need"),
        body("Choose all, answers only, inverse selection, or any custom mix of conversation fragments; edit and preview before you download.", 56, 158, 720, 48),
        rect(56, 250, 410, 176, "#124B49"),
        rect(490, 250, 390, 176, PALE),
        label("Selection scope", 78, 272, 160),
        text("All\nAnswers only\nInverse\nFree selection", 78, 310, 190, 100, { fontSize: 21, fontWeight: "bold", color: PAPER, lineSpacing: 1.05 }),
        label("Pre-delivery check", 512, 272, 180, INK),
        text("Edit content\nPreview result\nConfirm format", 512, 310, 190, 100, { fontSize: 21, fontWeight: "bold", color: INK, lineSpacing: 1.05 }),
        text("Less noise, clearer delivery boundaries.", 56, 456, 720, 28, { fontSize: 18, fontWeight: "bold", color: ORANGE }),
      ],
    },
    {
      id: "exportgpt-privacy",
      title: "Privacy policy",
      background: NAVY,
      elements: [
        ...base("08", "Privacy and processing"),
        title("Local-first processing keeps information movement bounded"),
        body("Most formats are processed locally in the browser; Screenshot and PDF use server-side processing, but the official site says it does not collect user data.", 56, 158, 760, 54),
        rect(56, 260, 240, 150, ORANGE),
        rect(360, 260, 240, 150, "#124B49"),
        rect(664, 260, 216, 150, CORAL),
        text("Local\nmost formats", 78, 300, 190, 65, { fontSize: 25, fontWeight: "bold", color: INK, align: "center", lineSpacing: 0.95 }),
        text("Server-side\nScreenshot / PDF", 382, 300, 196, 65, { fontSize: 22, fontWeight: "bold", color: PAPER, align: "center", lineSpacing: 0.95 }),
        text("Auto-delete\nwithin 24 hours", 686, 300, 172, 65, { fontSize: 22, fontWeight: "bold", color: PAPER, align: "center", lineSpacing: 0.95 }),
        text("The use case determines the right export path; clear processing boundaries make the product easier to trust long term.", 56, 456, 790, 28, { fontSize: 18, fontWeight: "bold", color: ORANGE }),
      ],
    },
    {
      id: "exportgpt-ecosystem",
      title: "Install and connect",
      background: NAVY,
      elements: [
        ...base("09", "Install and connect"),
        title("Install it beside ChatGPT and keep moving"),
        body("It installs in both Chrome and Edge; shared-link conversations can also be exported like normal chats.", 56, 158, 700, 45),
        roundRect(56, 255, 350, 132, ORANGE),
        roundRect(486, 255, 350, 132, CORAL),
        text("Chrome", 82, 278, 180, 30, { fontSize: 26, fontWeight: "bold", color: INK }),
        text("Open the Chrome Web Store\nand search for ExportGPT", 82, 322, 220, 46, { fontSize: 17, color: INK, lineSpacing: 1.1 }),
        text("Edge", 512, 278, 180, 30, { fontSize: 26, fontWeight: "bold", color: PAPER }),
        text("Open Microsoft Edge Add-ons\nand search for ExportGPT", 512, 322, 250, 46, { fontSize: 17, color: PAPER, lineSpacing: 1.1 }),
        text("No page switching, no habit changes, just a better exit path.", 56, 456, 720, 28, { fontSize: 18, fontWeight: "bold", color: ORANGE }),
      ],
    },
    {
      id: "exportgpt-close",
      title: "Close",
      background: DEEP,
      elements: [
        rect(0, 0, 20, H, ORANGE),
        rect(W - 160, 0, 80, H, ORANGE),
        rect(W - 80, 0, 80, H, CORAL),
        text("ExportGPT", 56, 78, 330, 34, { fontSize: 26, fontWeight: "bold", color: ORANGE }),
        text("Start with the next conversation\nand build your own AI knowledge base", 56, 155, 630, 115, { fontSize: 40, fontWeight: "bold", color: PAPER, lineSpacing: 0.95 }),
        line(56, 315, 145, ORANGE, 3),
        body("Export and use immediately.\nLet every strong answer be saved, shared, and reused.", 56, 342, 510, 68, PAPER),
        roundRect(56, 438, 230, 34, ORANGE),
        text("Install ExportGPT", 74, 446, 194, 18, { fontSize: 12, fontWeight: "bold", color: INK, align: "center" }),
        text("exportgpt.pro", 56, 510, 180, 16, { fontSize: 9, color: MUTED }),
        text("10", 855, 510, 50, 16, { fontSize: 9, color: MUTED, align: "right" }),
      ],
    },
  ],
};

export const exportgptProExample: ExampleDefinition = {
  id: "export-exportgpt-pro",
  title: "Export ExportGPT Product Case",
  feature: "export-pptx",
  description: "A ten-slide product introduction for exportgpt.pro using the supplied deck's visual language.",
  inputKind: "presentation-config",
  source: {
    label: "ExportGPT official website",
    content: "https://exportgpt.pro/\nOfficial product claims used: one-click export; 9 formats; selective export; edit and preview; image and share-link export; local-first processing; Screenshot/PDF server-side processing and automatic deletion within 24 hours.",
  },
  scenarioTags: ["product-introduction", "reference-style", "editable", "export"],
  expectedCapabilities: { normalize: "implemented", render: "implemented", exportPptx: "implemented" },
  status: "ready",
  createInput() {
    return structuredClone(input);
  },
};
