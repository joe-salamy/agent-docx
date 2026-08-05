import { readFile } from "node:fs/promises";
import { AgentDocxError } from "./types.js";

/**
 * One file-read entry point with a stable errno mapping: missing input is
 * `INPUT_NOT_FOUND` with the label; any other fs failure is `INTERNAL_ERROR`.
 */
export const readInputFile = async (
  path: string,
  label: string,
): Promise<Uint8Array> => {
  try {
    return await readFile(path);
  } catch (error) {
    const cause = error as NodeJS.ErrnoException;
    if (cause.code === "ENOENT" || cause.code === "ENOTDIR")
      throw new AgentDocxError("INPUT_NOT_FOUND", `${label} not found: ${path}`);
    throw new AgentDocxError(
      "INTERNAL_ERROR",
      error instanceof Error ? error.message : String(error),
    );
  }
};
