export type ThemeId = "clean-business" | "swiss-grid" | "editorial-story";

export type SlideRole =
  | "cover"
  | "agenda"
  | "section"
  | "statement"
  | "image"
  | "kpi"
  | "comparison"
  | "process"
  | "table"
  | "closing";

export interface SourceRef {
  id: string;
  label?: string;
}

export interface DeckBrief {
  title: string;
  audience: string;
  purpose: string;
  language: string;
  slideCountRange: [number, number];
  themeId: ThemeId;
  imagePolicy: string;
  constraints: string[];
  author?: string;
}

export interface ImagePlan {
  path: string;
  alt: string;
  width?: number;
  height?: number;
  fit?: "contain" | "cover" | "stretch";
}

export interface KpiPlan {
  value: string;
  label: string;
  detail?: string;
}

export interface ComparisonPlan {
  left: { heading: string; items: string[] };
  right: { heading: string; items: string[] };
}

export interface ChartPlan {
  type: "bar" | "line";
  categories: string[];
  series: Array<{ name: string; values: number[] }>;
}

export interface TablePlan {
  headers: string[];
  rows: string[][];
}

export interface SlidePlan {
  id: string;
  role: SlideRole;
  title: string;
  subtitle?: string;
  message?: string;
  items?: string[];
  image?: ImagePlan;
  kpis?: KpiPlan[];
  comparison?: ComparisonPlan;
  steps?: string[];
  table?: TablePlan;
  chart?: ChartPlan;
  notes?: string;
  sourceRefs?: SourceRef[];
}

export interface DeckSpec {
  brief: DeckBrief;
  slides: SlidePlan[];
}

export interface StructuralIssue {
  severity: "warning" | "error";
  code: string;
  message: string;
  slideId?: string;
  elementId?: string;
}

export interface BuildReport {
  output: string;
  slideCount: number;
  byteLength: number;
  diagnostics: Array<{ severity: string; code: string; message: string; path: string }>;
  exportWarnings: Array<{ code: string; message: string; slideId?: string; assetId?: string }>;
  structuralIssues: StructuralIssue[];
  packageChecks: { valid: boolean; parts: number; slideParts: number; issues: string[] };
  renderStatus: "not-run" | "skipped" | "partial" | "rendered";
  generatedAt: string;
}

export interface ExtractedSource {
  id: string;
  path: string;
  type: "text" | "document" | "table" | "image";
  content?: string;
  sheets?: Array<{ name: string; rows: unknown[][] }>;
  image?: { copiedPath: string; width?: number; height?: number };
  warnings: string[];
}
