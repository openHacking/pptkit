export type DiagnosticSeverity = "error" | "warning";

export interface PresentationDiagnostic {
  code: string;
  severity: DiagnosticSeverity;
  message: string;
  path: string;
  slideId?: string;
  elementId?: string;
  assetId?: string;
  layoutId?: string;
}

export class PresentationValidationError extends Error {
  readonly diagnostics: readonly PresentationDiagnostic[];

  constructor(diagnostics: readonly PresentationDiagnostic[]) {
    super(`Presentation validation failed with ${diagnostics.length} diagnostic${diagnostics.length === 1 ? "" : "s"}.`);
    this.name = "PresentationValidationError";
    this.diagnostics = diagnostics;
  }
}
