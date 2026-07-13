import { mkdir, stat, writeFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";

export async function writeOutput(path: string, bytes: Uint8Array): Promise<{ output: string; byteLength: number }> {
  const output = resolve(path);
  await mkdir(dirname(output), { recursive: true });
  await writeFile(output, bytes);
  return { output: path, byteLength: (await stat(output)).size };
}
