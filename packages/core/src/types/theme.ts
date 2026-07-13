import type { ThemeColorRole } from "./style.js";

export interface PresentationThemeInput {
  name?: string;
  colors?: Partial<Record<ThemeColorRole, string>>;
  fonts?: {
    heading?: string;
    body?: string;
  };
}

export interface NormalizedPresentationTheme {
  name: string;
  colors: Record<ThemeColorRole, string>;
  fonts: {
    heading: string;
    body: string;
  };
}
