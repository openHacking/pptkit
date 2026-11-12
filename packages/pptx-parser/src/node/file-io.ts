import * as fs from "node:fs/promises";
export type FileIo = Pick<typeof fs, "open" | "lstat" | "realpath" | "link" | "rename" | "unlink">;
export const nodeFileIo: FileIo = fs;
