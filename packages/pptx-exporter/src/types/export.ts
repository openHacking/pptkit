export interface ExportWarning {
  code: string;
  message: string;
  slideId?: string;
  elementIndex?: number;
  assetId?: string;
}

export interface GeneratePptxResult {
  bytes: Uint8Array;
  slideCount: number;
  byteLength: number;
  warnings: ExportWarning[];
  status: "generated" | "generated-with-warnings";
}

export interface WritePptxOptions {
  output: string;
}

export interface WritePptxResult {
  output: string;
  slideCount: number;
  byteLength: number;
  warnings: ExportWarning[];
  status: "written" | "written-with-warnings";
}
