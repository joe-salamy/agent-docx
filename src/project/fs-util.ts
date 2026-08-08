import { link, lstat, open, rename, rm } from "node:fs/promises";
import { randomUUID } from "node:crypto";
import { sep } from "node:path";

export const pathExists = async (path: string): Promise<boolean> => {
  try {
    await lstat(path);
    return true;
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === "ENOENT") return false;
    throw error;
  }
};

export const pathsOverlap = (left: string, right: string): boolean =>
  left === right ||
  left.startsWith(`${right}${sep}`) ||
  right.startsWith(`${left}${sep}`);

export const writeExclusiveFile = async (
  path: string,
  bytes: Uint8Array | string,
): Promise<void> => {
  const stage = `${path}.${randomUUID()}.stage`;
  let ownsStage = false;
  try {
    const handle = await open(stage, "wx", 0o600);
    ownsStage = true;
    try {
      await handle.writeFile(bytes);
      await handle.sync();
    } finally {
      await handle.close();
    }
    // Hard-link publication is atomic and, unlike rename, preserves wx
    // no-clobber semantics even when another writer races this call.
    await link(stage, path);
  } finally {
    if (ownsStage) await rm(stage, { force: true });
  }
};

/**
 * Persist bytes through a same-directory, fsynced stage before atomically
 * replacing the destination. Callers that need no-clobber semantics should
 * continue to use writeExclusiveFile for the destination itself.
 */
export const writeAtomicFile = async (
  path: string,
  bytes: Uint8Array | string,
): Promise<void> => {
  const stage = `${path}.${randomUUID()}.stage`;
  try {
    await writeExclusiveFile(stage, bytes);
    await rename(stage, path);
  } finally {
    await rm(stage, { force: true });
  }
};
