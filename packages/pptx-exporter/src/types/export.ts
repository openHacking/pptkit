export interface ExportPptxOptions {
  output: string;
}

export interface ExportWarning {
  code: string;
  message: string;
  slideId?: string;
  elementIndex?: number;
  assetId?: string;
}

export interface ExportResult {
  output: string;
  slideCount: number;
  byteLength: number;
  warnings: ExportWarning[];
  status: "written" | "written-with-warnings";
}
