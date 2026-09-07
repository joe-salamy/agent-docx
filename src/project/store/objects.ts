import { createHash, randomUUID } from "node:crypto";
import { constants as fsConstants } from "node:fs";
import {
  lstat,
  mkdir,
  open,
  readFile,
  readdir,
  realpath,
  rename,
  rm,
  type FileHandle,
} from "node:fs/promises";
import { dirname, parse, relative, resolve, sep } from "node:path";
import canonicalize from "canonicalize";
import { AgentDocxError } from "../../types.js";
import { writeAtomicFile, writeExclusiveFile } from "../fs-util.js";
import { isSafeRelativePath, publicPath } from "../../path-util.js";
import { builtInProfiles } from "../../profiles.js";
import { isDocumentId, type RevisionId } from "../../legal/model.js";
import { builtInRulePacks } from "../../legal/rules.js";
import type {
  AgentDocxDocumentConfig,
  DependencyHashes,
} from "../contracts.js";
import { EXPORT_INTENT, INIT_INTENT, LOCK_NAME, STORE_DIR } from "../store.js";
import type { OpenedStore, ProjectSnapshot } from "../store.js";

export const objectId = (bytes: Uint8Array | string): RevisionId =>
  `sha256:${createHash("sha256").update(bytes).digest("hex")}`;
export const emptyObjectId = objectId(new Uint8Array());

export const canonicalJson = (value: unknown): string => {
  const serialized = canonicalize(value);
  if (serialized === undefined)
    throw new AgentDocxError(
      "PROJECT_INVALID",
      "Value cannot be canonicalized",
    );
  return serialized;
};

export const relativePath = (
  projectDirectory: string,
  absolutePath: string,
): string => relative(projectDirectory, absolutePath).split(sep).join("/");

export const canonicalObjectId = (value: unknown): RevisionId =>
  objectId(canonicalJson(value));
const displayPath = (path: string, projectDirectory?: string): string =>
  projectDirectory === undefined ? path : publicPath(projectDirectory, path);

export const assertRelativeManifestPath = (
  path: string,
  name: string,
): string => {
  const parts = path.split("/");
  const reservedRoot =
    parts[0] === STORE_DIR ||
    parts[0] === INIT_INTENT ||
    parts[0] === EXPORT_INTENT ||
    parts[0] === LOCK_NAME ||
    parts[0]?.startsWith(".agent-docx.init-") === true;
  if (!isSafeRelativePath(path) || reservedRoot)
    throw new AgentDocxError(
      "PATH_OUTSIDE_PROJECT",
      `${name} must be a normalized project-relative path`,
    );
  return path;
};

export const assertNoSymlinkComponents = async (
  path: string,
  name: string,
): Promise<void> => {
  const absolutePath = resolve(path);
  const root = parse(absolutePath).root;
  let current = root;
  for (const component of relative(root, absolutePath).split(sep)) {
    if (component.length === 0) continue;
    current = resolve(current, component);
    let entry;
    try {
      entry = await lstat(current);
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code === "ENOENT") return;
      throw error;
    }
    if (entry.isSymbolicLink()) {
      // On macOS /var -> private/var and /tmp -> private/tmp are system symlinks;
      // allow those, but reject user-created symlinks.
      if (
        process.platform === "darwin" &&
        (current === "/var" || current === "/tmp")
      ) {
        try {
          const real = await realpath(current);
          if (
            (current === "/var" && real === "/private/var") ||
            (current === "/tmp" && real === "/private/tmp")
          ) {
            continue;
          }
        } catch {}
      }
      throw new AgentDocxError(
        "PATH_OUTSIDE_PROJECT",
        `${name} contains a symbolic-link component`,
      );
    }
  }
};

export const assertRegularFile = async (
  path: string,
  name: string,
  projectDirectory?: string,
): Promise<void> => {
  await assertNoSymlinkComponents(path, name);
  let entry;
  try {
    entry = await lstat(path);
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code !== "ENOENT") throw error;
    throw new AgentDocxError(
      "INPUT_NOT_FOUND",
      `${name} does not exist: ${displayPath(path, projectDirectory)}`,
    );
  }
  if (!entry.isFile() || entry.isSymbolicLink())
    throw new AgentDocxError(
      "PATH_OUTSIDE_PROJECT",
      `${name} must be a regular nonsymlink file: ${displayPath(path, projectDirectory)}`,
    );
};

export const assertDirectory = async (
  path: string,
  name: string,
  projectDirectory?: string,
): Promise<void> => {
  await assertNoSymlinkComponents(path, name);
  let entry;
  try {
    entry = await lstat(path);
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code !== "ENOENT") throw error;
    throw new AgentDocxError(
      "INPUT_NOT_FOUND",
      `${name} does not exist: ${displayPath(path, projectDirectory)}`,
    );
  }
  if (!entry.isDirectory() || entry.isSymbolicLink())
    throw new AgentDocxError(
      "PATH_OUTSIDE_PROJECT",
      `${name} must be a directory, not a symlink: ${displayPath(path, projectDirectory)}`,
    );
};

export const assertWithin = (
  projectDirectory: string,
  path: string,
  name: string,
): string => resolve(projectDirectory, assertRelativeManifestPath(path, name));

const sameFile = (
  left: { dev: number; ino: number },
  right: { dev: number; ino: number },
): boolean => left.dev === right.dev && left.ino === right.ino;

const regularFileError = (
  path: string,
  name: string,
  projectDirectory?: string,
): AgentDocxError =>
  new AgentDocxError(
    "PATH_OUTSIDE_PROJECT",
    `${name} must remain the same regular file: ${displayPath(path, projectDirectory)}`,
  );

/**
 * Read a project file through one stable file handle. The initial lstat and
 * final fstat pair catches replacement races; O_NOFOLLOW closes the final
 * symlink race on platforms that support it.
 */
export const readProjectFile = async (
  path: string,
  name: string,
  projectDirectory?: string,
): Promise<Uint8Array> => {
  await assertNoSymlinkComponents(path, name);
  let before;
  try {
    before = await lstat(path);
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code !== "ENOENT") throw error;
    throw new AgentDocxError(
      "INPUT_NOT_FOUND",
      `${name} does not exist: ${displayPath(path, projectDirectory)}`,
    );
  }
  if (!before.isFile() || before.isSymbolicLink())
    throw regularFileError(path, name, projectDirectory);
  const noFollow = fsConstants.O_NOFOLLOW ?? 0;
  let handle: FileHandle | undefined;
  try {
    try {
      handle = await open(path, fsConstants.O_RDONLY | noFollow);
    } catch (error) {
      const code = (error as NodeJS.ErrnoException).code;
      if (
        noFollow === 0 ||
        !["EINVAL", "ENOTSUP", "EOPNOTSUPP"].includes(code ?? "")
      )
        throw error;
      handle = await open(path, fsConstants.O_RDONLY);
    }
    const opened = await handle.stat();
    if (
      !opened.isFile() ||
      opened.isSymbolicLink() ||
      !sameFile(before, opened)
    )
      throw regularFileError(path, name, projectDirectory);
    const bytes = await handle.readFile();
    const after = await handle.stat();
    if (!after.isFile() || after.isSymbolicLink() || !sameFile(opened, after))
      throw regularFileError(path, name, projectDirectory);
    return bytes;
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === "ELOOP")
      throw regularFileError(path, name, projectDirectory);
    throw error;
  } finally {
    await handle?.close();
  }
};

export const strictJson = <T>(bytes: Uint8Array, path: string): T => {
  try {
    return JSON.parse(
      new TextDecoder("utf-8", { fatal: true }).decode(bytes),
    ) as T;
  } catch (error) {
    throw new AgentDocxError("PROJECT_INVALID", `Invalid JSON: ${path}`, {
      cause: error instanceof Error ? error.message : String(error),
    });
  }
};

export const readJsonFile = async <T>(path: string): Promise<T> =>
  strictJson<T>(await readFile(path), path);

const objectPath = (storePath: string, id: RevisionId): string => {
  const hex = id.slice("sha256:".length);
  if (!/^[0-9a-f]{64}$/.test(hex))
    throw new AgentDocxError("PROJECT_INVALID", `Invalid object id: ${id}`);
  return resolve(storePath, "objects", "sha256", hex.slice(0, 2), hex.slice(2));
};

export const writeObject = async (
  storePath: string,
  bytes: Uint8Array | string,
): Promise<RevisionId> => {
  const id = objectId(bytes);
  const path = objectPath(storePath, id);
  await mkdir(dirname(path), { recursive: true, mode: 0o700 });
  try {
    await writeAtomicFile(path, bytes);
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code !== "EEXIST") throw error;
    const current = await readFile(path);
    if (objectId(current) !== id)
      throw new AgentDocxError("PROJECT_INVALID", `Object collision at ${id}`);
  }
  return id;
};

export const readObject = async (
  storePath: string,
  id: RevisionId,
): Promise<Uint8Array> => {
  const path = objectPath(storePath, id);
  await assertRegularFile(path, "Object");
  const bytes = await readFile(path);
  if (objectId(bytes) !== id)
    throw new AgentDocxError("PROJECT_INVALID", `Object hash mismatch: ${id}`);
  return bytes;
};

const refsPath = (storePath: string, documentId: string): string => {
  if (!isDocumentId(documentId))
    throw new AgentDocxError(
      "DOCUMENT_NOT_FOUND",
      `Invalid document id: ${documentId}`,
    );
  return resolve(storePath, "refs", `${documentId}.json`);
};

const revisionPath = (storePath: string, id: RevisionId): string => {
  const hex = id.slice("sha256:".length);
  if (!/^[0-9a-f]{64}$/.test(hex))
    throw new AgentDocxError(
      "REVISION_NOT_FOUND",
      `Invalid revision id: ${id}`,
    );
  return resolve(storePath, "revisions", `${hex}.json`);
};

export const readHead = async (
  storePath: string,
  documentId: string,
): Promise<RevisionId | null> => {
  const path = refsPath(storePath, documentId);
  try {
    const entry = await readJsonFile<{
      schemaVersion: 1;
      documentId: string;
      head: RevisionId | null;
    }>(path);
    if (entry.schemaVersion !== 1 || entry.documentId !== documentId)
      throw new AgentDocxError(
        "PROJECT_INVALID",
        `Malformed head reference: ${documentId}`,
      );
    if (entry.head !== null && !/^sha256:[0-9a-f]{64}$/.test(entry.head))
      throw new AgentDocxError(
        "PROJECT_INVALID",
        `Malformed head reference: ${documentId}`,
      );
    return entry.head;
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === "ENOENT") return null;
    throw error;
  }
};

export const writeHead = async (
  storePath: string,
  documentId: string,
  head: RevisionId,
): Promise<void> => {
  const path = refsPath(storePath, documentId);
  await mkdir(dirname(path), { recursive: true, mode: 0o700 });
  const stage = `${path}.${randomUUID()}.stage`;
  await writeExclusiveFile(
    stage,
    canonicalJson({ schemaVersion: 1, documentId, head }),
  );
  try {
    await rename(stage, path);
  } finally {
    await rm(stage, { force: true });
  }
};

export const readRevisionJson = async <T>(
  storePath: string,
  revision: RevisionId,
): Promise<T> => {
  const path = revisionPath(storePath, revision);
  const record = strictJson<Record<string, unknown>>(
    await readFile(path),
    path,
  );
  const withoutId = { ...record };
  delete withoutId.id;
  if (record.id !== revision || canonicalObjectId(withoutId) !== revision)
    throw new AgentDocxError(
      "PROJECT_INVALID",
      `Revision hash mismatch: ${revision}`,
    );
  return record as T;
};

export const writeRevisionJson = async (
  storePath: string,
  value: Record<string, unknown>,
): Promise<RevisionId> => {
  const withoutId = { ...value };
  delete withoutId.id;
  const id = canonicalObjectId(withoutId);
  const record = canonicalJson({ ...withoutId, id });
  const path = revisionPath(storePath, id);
  await mkdir(dirname(path), { recursive: true, mode: 0o700 });
  try {
    await writeAtomicFile(path, record);
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code !== "EEXIST") throw error;
    const existing = strictJson<Record<string, unknown>>(
      await readFile(path),
      path,
    );
    const existingWithoutId = { ...existing };
    delete existingWithoutId.id;
    if (existing.id !== id || canonicalObjectId(existingWithoutId) !== id)
      throw new AgentDocxError(
        "PROJECT_INVALID",
        `Revision collision at ${id}`,
      );
  }
  return id;
};

const mediaTypeFor = (path: string): string => {
  const lower = path.toLowerCase();
  if (lower.endsWith(".png")) return "image/png";
  if (lower.endsWith(".jpg") || lower.endsWith(".jpeg")) return "image/jpeg";
  if (lower.endsWith(".pdf")) return "application/pdf";
  return "application/octet-stream";
};

const collectAssets = async (
  directory: string,
  relativeDirectory = "",
  projectDirectory?: string,
): Promise<ReadonlyMap<string, { bytes: Uint8Array; mediaType: string }>> => {
  const assets = new Map<string, { bytes: Uint8Array; mediaType: string }>();
  const entries = await readdir(directory, { withFileTypes: true });
  for (const entry of entries.sort((left, right) =>
    left.name < right.name ? -1 : left.name > right.name ? 1 : 0,
  )) {
    const logical = relativeDirectory
      ? `${relativeDirectory}/${entry.name}`
      : entry.name;
    const path = resolve(directory, entry.name);
    if (entry.isSymbolicLink())
      throw new AgentDocxError(
        "PATH_OUTSIDE_PROJECT",
        `Asset is a symlink: ${logical}`,
      );
    if (entry.isDirectory()) {
      for (const [key, value] of await collectAssets(
        path,
        logical,
        projectDirectory,
      ))
        assets.set(key, value);
      continue;
    }
    if (!entry.isFile())
      throw new AgentDocxError(
        "REFERENCE_INVALID",
        `Unsupported asset entry: ${logical}`,
      );
    assets.set(logical, {
      bytes: await readProjectFile(path, `Asset ${logical}`, projectDirectory),
      mediaType: mediaTypeFor(logical),
    });
  }
  return assets;
};

const addDependency = async (
  bytes: Uint8Array | string,
  mediaType: string,
  key: string,
  dependencies: Map<string, { bytes: Uint8Array; mediaType: string }>,
): Promise<RevisionId> => {
  const encoded =
    typeof bytes === "string" ? new TextEncoder().encode(bytes) : bytes;
  dependencies.set(key, { bytes: encoded, mediaType });
  return objectId(encoded);
};

export const snapshotProjectDocument = async (
  opened: OpenedStore,
  document: AgentDocxDocumentConfig,
): Promise<ProjectSnapshot> => {
  const sourcePath = assertWithin(
    opened.projectDirectory,
    document.source,
    "Document source",
  );
  const sourceBytes = await readProjectFile(
    sourcePath,
    "Document source",
    opened.projectDirectory,
  );
  let source: string;
  try {
    source = new TextDecoder("utf-8", { fatal: true }).decode(sourceBytes);
  } catch {
    throw new AgentDocxError(
      "INPUT_NOT_UTF8",
      `Document source is not UTF-8: ${document.source}`,
    );
  }
  const dependencies = new Map<
    string,
    { bytes: Uint8Array; mediaType: string }
  >();
  const dependencyObjects: Record<string, RevisionId> = {};
  const profile = builtInProfiles[document.profile];
  if (!profile)
    throw new AgentDocxError(
      "PROJECT_INVALID",
      `Unknown profile: ${document.profile}`,
    );
  dependencyObjects.profile = await addDependency(
    canonicalJson(profile),
    "application/json",
    "profile",
    dependencies,
  );
  if (document.rulePack) {
    const pack = builtInRulePacks[document.rulePack];
    if (!pack)
      throw new AgentDocxError(
        "PROJECT_INVALID",
        `Unknown rule pack: ${document.rulePack}`,
      );
    dependencyObjects["rule-pack"] = await addDependency(
      canonicalJson(pack),
      "application/json",
      "rule-pack",
      dependencies,
    );
    const source = await readFile(
      new URL(`../../../assets/rules/${pack.sourceExcerpt}`, import.meta.url),
    );
    if (objectId(source) !== pack.sourceSha256)
      throw new AgentDocxError(
        "PROJECT_INVALID",
        `Rule source hash mismatch: ${pack.sourceExcerpt}`,
      );
    dependencyObjects[`rule-source/${pack.sourceExcerpt}`] =
      await addDependency(
        source,
        "text/plain",
        `rule-source/${pack.sourceExcerpt}`,
        dependencies,
      );
  }
  for (const [index, configuredPath] of (document.rulePacks ?? []).entries()) {
    const packPath = assertWithin(
      opened.projectDirectory,
      configuredPath,
      `Rule pack ${index}`,
    );
    const bytes = await readProjectFile(
      packPath,
      `Rule pack ${index}`,
      opened.projectDirectory,
    );
    dependencyObjects[`rule-pack:${index}`] = await addDependency(
      bytes,
      "application/json",
      `rule-pack:${index}`,
      dependencies,
    );
  }
  if (document.template) {
    const templatePath = assertWithin(
      opened.projectDirectory,
      document.template,
      "Template",
    );
    const bytes = await readProjectFile(
      templatePath,
      "Template",
      opened.projectDirectory,
    );
    dependencyObjects.template = await addDependency(
      bytes,
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      "template",
      dependencies,
    );
  }
  if (document.assetsDir) {
    const assetsPath = assertWithin(
      opened.projectDirectory,
      document.assetsDir,
      "Assets directory",
    );
    await assertDirectory(
      assetsPath,
      "Assets directory",
      opened.projectDirectory,
    );
    for (const [name, asset] of await collectAssets(
      assetsPath,
      "",
      opened.projectDirectory,
    ))
      dependencyObjects[`asset/${name}`] = await addDependency(
        asset.bytes,
        asset.mediaType,
        `asset/${name}`,
        dependencies,
      );
  }
  if (document.fontSet) {
    const roles: Readonly<Record<string, string | undefined>> = {
      regular: document.fontSet.regularPath,
      bold: document.fontSet.boldPath,
      italic: document.fontSet.italicPath,
      boldItalic: document.fontSet.boldItalicPath,
    };
    for (const [role, configuredPath] of Object.entries(roles)) {
      if (!configuredPath) continue;
      const fontPath = assertWithin(
        opened.projectDirectory,
        configuredPath,
        `Font ${role}`,
      );
      const bytes = await readProjectFile(
        fontPath,
        `Font ${role}`,
        opened.projectDirectory,
      );
      dependencyObjects[`font/${role}`] = await addDependency(
        bytes,
        "font/ttf",
        `font/${role}`,
        dependencies,
      );
    }
  }
  const sourceObject = objectId(sourceBytes);
  const documentConfigObject = canonicalObjectId(document);
  const sortedDependencies = Object.fromEntries(
    Object.entries(dependencyObjects).sort(([left], [right]) =>
      left < right ? -1 : left > right ? 1 : 0,
    ),
  ) as DependencyHashes;
  const workingTreeHash = canonicalObjectId({
    sourceObject,
    documentConfigObject,
    dependencyObjects: sortedDependencies,
  });
  return {
    source,
    sourceObject,
    documentConfigObject,
    dependencyObjects: sortedDependencies,
    dependencyBytes: dependencies,
    workingTreeHash,
  };
};

export const storeSnapshot = async (
  opened: OpenedStore,
  snapshot: ProjectSnapshot,
): Promise<void> => {
  await writeObject(opened.storePath, snapshot.source);
  await writeObject(opened.storePath, canonicalJson(opened.manifest));
  for (const { bytes } of snapshot.dependencyBytes.values())
    await writeObject(opened.storePath, bytes);
};
