import { runCliAsync, type CliRunResult } from "@pptkit/cli";

export async function inspectLocalDeck(file: string, signal?: AbortSignal): Promise<CliRunResult> {
  return runCliAsync(["inspect", file, "--projection", "content", "--json"], signal ? { signal } : {});
}
