import type { DeckSpec } from "./contracts.js";

export const deckSpec: DeckSpec = {
  brief: {
    title: "Untitled PPTKit Deck",
    audience: "Project stakeholders",
    purpose: "Explain one clear idea and agree on the next action",
    language: "en-US",
    slideCountRange: [3, 5],
    themeId: "clean-business",
    imagePolicy: "Use only supplied local assets",
    constraints: ["Keep every visible object editable in PowerPoint"],
    author: "PPTKit",
  },
  slides: [
    {
      id: "cover",
      role: "cover",
      title: "Replace with the confirmed title",
      subtitle: "Replace with the audience promise",
    },
    {
      id: "core-message",
      role: "statement",
      title: "The core message",
      message: "Replace this placeholder with one defensible, source-backed claim.",
      sourceRefs: [],
    },
    {
      id: "closing",
      role: "closing",
      title: "Next step",
      message: "Replace with the action the audience should take.",
    },
  ],
};
