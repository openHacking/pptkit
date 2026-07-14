import type { ThemeId } from "../contracts.js";

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
  titleSize: number;
  bodySize: number;
  captionSize: number;
  margin: number;
  gap: number;
  radius: number;
}

export const themes: Record<ThemeId, ThemeTokens> = {
  "clean-business": {
    id: "clean-business",
    name: "PPTKit Clean Business",
    background: "F7F9FC",
    surface: "FFFFFF",
    text: "10243E",
    muted: "617086",
    accent: "2457D6",
    accent2: "16A394",
    headingFont: "Arial",
    bodyFont: "Arial",
    titleSize: 34,
    bodySize: 19,
    captionSize: 11,
    margin: 56,
    gap: 20,
    radius: 12,
  },
  "swiss-grid": {
    id: "swiss-grid",
    name: "PPTKit Swiss Grid",
    background: "F5F4EF",
    surface: "FFFFFF",
    text: "111111",
    muted: "5F5F5F",
    accent: "E33B2E",
    accent2: "1D4ED8",
    headingFont: "Arial",
    bodyFont: "Arial",
    titleSize: 38,
    bodySize: 18,
    captionSize: 11,
    margin: 48,
    gap: 16,
    radius: 0,
  },
  "editorial-story": {
    id: "editorial-story",
    name: "PPTKit Editorial Story",
    background: "F3EEE5",
    surface: "FBF8F2",
    text: "2B2926",
    muted: "746D64",
    accent: "B94F35",
    accent2: "2F6D62",
    headingFont: "Georgia",
    bodyFont: "Arial",
    titleSize: 36,
    bodySize: 18,
    captionSize: 11,
    margin: 60,
    gap: 22,
    radius: 4,
  },
};

export function getTheme(themeId: ThemeId): ThemeTokens {
  return themes[themeId];
}
