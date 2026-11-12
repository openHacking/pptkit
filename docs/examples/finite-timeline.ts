import { compileFiniteTimeline, sampleFiniteTimeline, type FiniteTimelineNode } from "@pptkit/core";

export function finiteTimingExample(timeMs: number) {
  const timing: FiniteTimelineNode = {
    kind: "sequence", id: "sequence", delayMs: 50,
    children: [
      { kind: "span", id: "first", durationMs: 100, repeat: 2, fill: "hold" },
      { kind: "parallel", id: "parallel", children: [
        { kind: "span", id: "second", durationMs: 80 },
        { kind: "span", id: "instant", delayMs: 20, durationMs: 0, fill: "hold" },
      ] },
    ],
  };
  const plan = compileFiniteTimeline(timing);
  // Keep the authored tree for editing. A JSON schedule can be transported
  // and sampled, but it does not encode drawing targets or native effects.
  const transported = JSON.parse(JSON.stringify(plan)) as typeof plan;
  const samples = sampleFiniteTimeline(transported, timeMs);
  return { plan, samples };
}
