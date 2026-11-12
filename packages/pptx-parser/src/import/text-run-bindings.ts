import type { ImportedNotesTextRunSource, ImportedTableTextRunSource, ImportedTextRunSource } from "../types/import.js";
import type { structureContext } from "../presentation/structure-context.js";
import type { ContentTextRun } from "./content-text.js";

export type TextRunBinding = (ImportedTextRunSource | (ImportedNotesTextRunSource & { readonly owner: "notes"; readonly elementId?: never }) | (ImportedTableTextRunSource & { readonly owner: "table-cell" })) & {
  readonly baselineText: string;
  readonly namespaceUri: string;
  /** Direct raw a:r/a:p ancestry only; evidence, not edit authority. */
  readonly rawRunElementIndex?: number;
  readonly rawParagraphElementIndex?: number;
};
type TextRunTarget = { readonly slideId: string; readonly elementId: string; readonly owner?: "text-box" | "shape-text" } | { readonly slideId: string; readonly owner: "notes" } | { readonly slideId: string; readonly elementId: string; readonly owner: "table-cell"; readonly rowIndex: number; readonly cellIndex: number; readonly columnIndex: number };

export async function bindTextRuns(
  runs: readonly ContentTextRun[], target: TextRunTarget,
  source: ImportedTextRunSource["source"], context: Awaited<ReturnType<typeof structureContext>>,
  checkpoint: () => Promise<void>,
): Promise<TextRunBinding[]> {
  const result: TextRunBinding[] = [];
  for (const run of runs) {
    await checkpoint();
    const raw = run.node ? await context.sourceElement(source.part, run.node) : undefined;
    const characterOnly = raw && raw.node.children.every(child => child.kind === "text" || child.kind === "cdata");
    const value = characterOnly ? raw.node.children.map(child => child.kind === "text" || child.kind === "cdata" ? child.value : "").join("") : undefined;
    const bound = Boolean(run.editable !== false && raw && characterOnly && value === run.text);
    const rawRun = raw?.parent?.namespaceUri === context.a && raw.parent.local === "r" ? await context.sourceElement(source.part, raw.parent) : undefined;
    const rawParagraph = rawRun?.parent?.namespaceUri === context.a && rawRun.parent.local === "p" ? await context.sourceElement(source.part, rawRun.parent) : undefined;
    result.push({
      ...target, paragraphIndex: run.paragraphIndex, runIndex: run.runIndex,
      source: { ...source, ...(raw ? { elementIndex: raw.elementIndex } : {}) }, baselineText: run.text, namespaceUri: context.a,
      ...(raw ? { sourceElementIndex: raw.elementIndex } : {}),
      ...(rawRun ? { rawRunElementIndex: rawRun.elementIndex } : {}),
      ...(rawParagraph ? { rawParagraphElementIndex: rawParagraph.elementIndex } : {}),
      status: bound ? "bound" : "unmapped",
      ...(bound ? {} : { reason: run.editable === false ? "Dynamic field caches are not editable through scalar text preservation." : raw ? "Original character content is mixed or differs from the effective projection." : "Break/synthetic insertion runs have no direct source text element." }),
    });
  }
  return result;
}
