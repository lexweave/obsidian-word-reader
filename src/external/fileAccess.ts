import type { Stats } from "node:fs";
import type * as fs from "node:fs/promises";

export const MAX_EXTERNAL_PPTX_BYTES = 512 * 1024 * 1024;

export function revision(stat: Stats): string {
  return [stat.dev, stat.ino, stat.size, stat.mtimeMs, stat.ctimeMs].join(":");
}

export function validateStat(stat: Stats): void {
  if (!stat.isFile() || stat.size <= 0 || stat.size > MAX_EXTERNAL_PPTX_BYTES) {
    throw new Error("The presentation is not a regular file or exceeds 512 MiB.");
  }
}

/** Read one descriptor and reject replacements/truncation instead of mixing versions. */
export async function readPresentation(
  io: Pick<typeof fs, "open" | "stat">,
  path: string,
  expected: string,
  flags: string | number = "r",
): Promise<ArrayBuffer> {
  const file = await io.open(path, flags);
  try {
    const before = await file.stat();
    validateStat(before);
    if (revision(before) !== expected) throw new Error("The presentation changed. Reload it.");
    // Bounded reads avoid readFile allocating beyond the limit if a file grows.
    const data = new Uint8Array(before.size);
    let offset = 0;
    while (offset < data.length) {
      const { bytesRead } = await file.read(data, offset, Math.min(1024 * 1024, data.length - offset), offset);
      if (!bytesRead) throw new Error("The presentation was truncated. Reload it.");
      offset += bytesRead;
    }
    const after = await file.stat();
    const current = await io.stat(path);
    if (revision(before) !== revision(after) || revision(before) !== revision(current)) {
      throw new Error("The presentation changed while it was being read. Reload it.");
    }
    return data.buffer;
  } finally { await file.close(); }
}

export function isWithinRoot(relative: string, separator: string): boolean {
  return relative === "" || (!relative.startsWith(".." + separator) && relative !== ".." &&
    !relative.startsWith(separator) && !/^[a-z]:/i.test(relative));
}
