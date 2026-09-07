import { realpath } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { lock } from "proper-lockfile";
import { AgentDocxError } from "../../types.js";
import { publicPath } from "../../path-util.js";
import { LOCK_NAME, STORE_DIR, type OpenedStore } from "../store.js";
import {
  assertDirectory,
  assertRegularFile,
  readProjectFile,
  recoverExport,
  recoverInitialization,
  strictJson,
  sweepOwnedStages,
  validateBinding,
  validateManifest,
  validateManifestPaths,
} from "./recovery.js";

export const acquireProjectLock = async (
  projectDirectory: string,
): Promise<() => Promise<void>> => {
  const canonicalDirectory = await realpath(projectDirectory);
  const lockPath = resolve(canonicalDirectory, LOCK_NAME);
  try {
    return await lock(canonicalDirectory, {
      realpath: true,
      lockfilePath: lockPath,
      retries: 0,
      // A blocked event loop can delay heartbeat updates; this bounded
      // window still trades rare false stale-lock detection for safety.
      stale: 60000,
      update: 15000,
    });
  } catch (error) {
    throw new AgentDocxError("PROJECT_LOCKED", "Project is locked", {
      cause: error instanceof Error ? error.message : String(error),
    });
  }
};

export const openStore = (manifestPath: string): Promise<OpenedStore> =>
  withLockedStore<OpenedStore>(manifestPath, async (opened) => opened);

export const withLockedStore = async <Value>(
  manifestPath: string,
  operation: (opened: OpenedStore) => Promise<Value>,
): Promise<Value> => {
  const absoluteManifestPath = resolve(manifestPath);
  const projectDirectory = dirname(absoluteManifestPath);
  await assertDirectory(
    projectDirectory,
    "Project directory",
    projectDirectory,
  );
  const release = await acquireProjectLock(projectDirectory);
  let opening = true;
  try {
    await recoverInitialization(projectDirectory, absoluteManifestPath);
    await recoverExport(projectDirectory, absoluteManifestPath);
    await sweepOwnedStages(projectDirectory, absoluteManifestPath);
    await assertRegularFile(
      absoluteManifestPath,
      "Project manifest",
      projectDirectory,
    );
    const manifest = validateManifest(
      strictJson<unknown>(
        await readProjectFile(
          absoluteManifestPath,
          "Project manifest",
          projectDirectory,
        ),
        absoluteManifestPath,
      ),
    );
    await validateManifestPaths(projectDirectory, manifest);
    const opened = {
      manifestPath: absoluteManifestPath,
      projectDirectory: await realpath(projectDirectory),
      storePath: resolve(projectDirectory, STORE_DIR),
      manifest,
    };
    await assertDirectory(
      opened.storePath,
      "Project store",
      opened.projectDirectory,
    );
    await validateBinding(opened);
    opening = false;
    return await operation(opened);
  } catch (error) {
    const code =
      error instanceof AgentDocxError
        ? error.code
        : (error as NodeJS.ErrnoException).code;
    if (opening && (code === "INPUT_NOT_FOUND" || code === "ENOENT"))
      throw new AgentDocxError(
        "PROJECT_NOT_FOUND",
        `Project not found: ${publicPath(projectDirectory, absoluteManifestPath)}`,
      );
    throw error;
  } finally {
    await release();
  }
};
