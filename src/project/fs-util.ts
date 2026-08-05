import { lstat, open } from "node:fs/promises";
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
  const handle = await open(path, "wx", 0o600);
  try {
    await handle.writeFile(bytes);
  } finally {
    await handle.close();
  }
};
