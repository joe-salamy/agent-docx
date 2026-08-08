import { constants as fsConstants } from "node:fs";
import { open, type FileHandle } from "node:fs/promises";
import { AgentDocxError } from "./types.js";

export const MAX_INPUT_BYTES = 64 * 1024 * 1024;
export const MAX_FONT_BYTES = 32 * 1024 * 1024;

const inputTooLarge = (label: string, limit: number): AgentDocxError =>
  new AgentDocxError(
    "INPUT_TOO_LARGE",
    `${label} exceeds the ${limit} byte input limit`,
  );

const CHUNK_BYTES = 64 * 1024;

/**
 * Reads a file through one stable handle with a bounded loop so a file that
 * grows or is swapped after an initial size check can never bypass the byte
 * cap (the stat-then-read pattern is a TOCTOU window). Non-regular files and
 * symlinks are rejected before any bytes are consumed.
 */
export const readInputFile = async (
  path: string,
  label: string,
  maxBytes = MAX_INPUT_BYTES,
): Promise<Uint8Array> => {
  let handle: FileHandle | undefined;
  try {
    const flags =
      process.platform === "win32"
        ? "r"
        : fsConstants.O_RDONLY | fsConstants.O_NOFOLLOW;
    handle = await open(path, flags);
    const information = await handle.stat();
    if (!information.isFile() || information.isSymbolicLink())
      throw new AgentDocxError("INPUT_NOT_FOUND", `${label} not found`);
    if (information.size > maxBytes) throw inputTooLarge(label, maxBytes);
    const chunks: Uint8Array[] = [];
    let total = 0;
    const buffer = new Uint8Array(CHUNK_BYTES);
    while (true) {
      const { bytesRead } = await handle.read(buffer, 0, CHUNK_BYTES, null);
      if (bytesRead === 0) break;
      total += bytesRead;
      if (total > maxBytes) throw inputTooLarge(label, maxBytes);
      chunks.push(bytesRead === CHUNK_BYTES ? buffer.slice() : buffer.slice(0, bytesRead));
    }
    if (chunks.length === 1) return chunks[0]!;
    const output = new Uint8Array(total);
    let offset = 0;
    for (const chunk of chunks) {
      output.set(chunk, offset);
      offset += chunk.byteLength;
    }
    return output;
  } catch (error) {
    if (error instanceof AgentDocxError) throw error;
    const cause = error as NodeJS.ErrnoException;
    if (
      cause.code === "ENOENT" ||
      cause.code === "ENOTDIR" ||
      cause.code === "ELOOP"
    )
      throw new AgentDocxError("INPUT_NOT_FOUND", `${label} not found`);
    throw new AgentDocxError("INTERNAL_ERROR", `${label} could not be read`);
  } finally {
    await handle?.close();
  }
};
