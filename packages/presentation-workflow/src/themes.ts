import type { ThemeId } from "./contracts.js";

export const SLIDE = { width: 960, height: 540 } as const;

export interface ThemeTokens {
  id: ThemeId;
  name: string;
  background: string;
  surface: string;
  text: string;
  muted: string;
  accent: string;
  accent2: string;
  headingFont: string;
  bodyFont: string;
  margin: number;
  gap: number;
  radius: number;
  titleSize: number;
  bodySize: number;
  captionSize: number;
}

const THEMES: Record<ThemeId, ThemeTokens> = {
  "clean-business": {
    id: "clean-business", name: "Clean Business", background: "F7F8FA", surface: "FFFFFF",
    text: "172033", muted: "647083", accent: "2457D6", accent2: "11A683",
    headingFont: "Aptos Display", bodyFont: "Aptos", margin: 56, gap: 20, radius: 14,
    titleSize: 30, bodySize: 17, captionSize: 11,
  },
  "swiss-grid": {
    id: "swiss-grid", name: "Swiss Grid", background: "F4F1EA", surface: "FFFFFF",
    text: "111111", muted: "5F5A50", accent: "E6402B", accent2: "2255CC",
    headingFont: "Arial", bodyFont: "Arial", margin: 48, gap: 16, radius: 0,
    titleSize: 32, bodySize: 17, captionSize: 11,
  },
  "editorial-story": {
    id: "editorial-story", name: "Editorial Story", background: "FBF7F0", surface: "FFFDFC",
    text: "2B2118", muted: "766A5E", accent: "A33C2E", accent2: "236A73",
    headingFont: "Georgia", bodyFont: "Aptos", margin: 60, gap: 24, radius: 8,
    titleSize: 31, bodySize: 18, captionSize: 11,
  },
};

export function getTheme(id: ThemeId): ThemeTokens {
  return THEMES[id];
}
