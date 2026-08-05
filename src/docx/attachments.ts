import { lstat, readdir } from "node:fs/promises";
import { basename, isAbsolute, relative, resolve, sep } from "node:path";
import { AgentDocxError } from "../types.js";
import type { SemanticManifest } from "./manifest.js";
import type {
  AttachmentManifest,
  ImportAttachmentBundle,
} from "./contracts.js";
import {
  decodeDocxXml,
  sha256Hex,
} from "./package.js";
import { readInputFile } from "../input.js";
const unsupported = (message: string): never => {
  throw new AgentDocxError("DOCX_IMPORT_UNSUPPORTED", message);
};

const asObject = (value: unknown, label: string): Record<string, unknown> => {
  if (value === null || typeof value !== "object" || Array.isArray(value))
    unsupported(`${label} must be an object`);
  return value as Record<string, unknown>;
};

const exactKeys = (
  value: Record<string, unknown>,
  keys: readonly string[],
  label: string,
): void => {
  if (
    Object.keys(value).length !== keys.length ||
    Object.keys(value).some((key) => !keys.includes(key))
  )
    unsupported(`${label} has an unsupported property`);
};



export const attachmentInventory = (
  parts: ReadonlyMap<string, Uint8Array>,
): Record<
  string,
  { sha256: `sha256:${string}`; mediaType: string; bytes: number }
> => {
  const assets: Record<
    string,
    { sha256: `sha256:${string}`; mediaType: string; bytes: number }
  > = {};
  for (const [path, bytes] of parts) {
    if (!path.startsWith("word/media/")) continue;
    const name = basename(path);
    const mediaType = path.endsWith(".png")
      ? "image/png"
      : /\.jpe?g$/i.test(path)
        ? "image/jpeg"
        : "application/octet-stream";
    assets[name] = {
      sha256: sha256Hex(bytes),
      mediaType,
      bytes: bytes.byteLength,
    };
  }
  return assets;
};

export type ImportedAsset = { bytes: Uint8Array; mediaType: string };

export type AttachmentResolution = {
  assets: Readonly<Record<string, ImportedAsset>>;
  inventory: Readonly<
    Record<
      string,
      { sha256: `sha256:${string}`; mediaType: string; bytes: number }
    >
  >;
  complete: boolean;
};

const sameAttachmentEntries = (
  left: readonly AttachmentManifest["entries"][number][],
  right: readonly AttachmentManifest["entries"][number][],
): boolean =>
  left.length === right.length &&
  left.every(
    (entry, index) =>
      entry.name === right[index]!.name &&
      entry.mediaType === right[index]!.mediaType &&
      entry.byteLength === right[index]!.byteLength &&
      entry.sha256 === right[index]!.sha256 &&
      entry.payloadPath === right[index]!.payloadPath,
  );

const pathInside = (root: string, path: string, label: string): string => {
  const target = resolve(root, path);
  const contained = relative(root, target);
  if (
    contained.length === 0 ||
    isAbsolute(contained) ||
    contained.split(sep).some((part) => part === "..")
  )
    unsupported(`${label} escapes its attachment bundle`);
  return target;
};

const ATTACHMENT_MAX_ENTRIES = 512;
const ATTACHMENT_MAX_FILE_BYTES = 25 * 1024 * 1024;
const ATTACHMENT_MAX_TOTAL_BYTES = 50 * 1024 * 1024;

const regularAttachmentFile = async (
  path: string,
  label: string,
): Promise<Uint8Array> => {
  const entry = await lstat(path).catch(() => null);
  if (!entry || !entry.isFile() || entry.isSymbolicLink())
    unsupported(`${label} is not a regular nonsymlink file`);
  const stats = entry as NonNullable<typeof entry>;
  if (stats.size > ATTACHMENT_MAX_FILE_BYTES)
    throw new AgentDocxError(
      "DOCX_TOO_LARGE",
      `${label} exceeds ${ATTACHMENT_MAX_FILE_BYTES / (1024 * 1024)} MiB`,
    );
  return readInputFile(path, label);
};

const bundleFiles = async (directory: string): Promise<readonly string[]> => {
  const root = await lstat(directory).catch(() => null);
  if (!root || !root.isDirectory() || root.isSymbolicLink())
    unsupported("Attachment bundle directory is not a real directory");
  let fileCount = 0;
  let totalBytes = 0;
  const visit = async (path: string, prefix: string): Promise<string[]> => {
    const entries = await readdir(path, { withFileTypes: true });
    const files: string[] = [];
    for (const entry of entries) {
      if (entry.isSymbolicLink())
        unsupported("Attachment bundle contains a symbolic link");
      const relativePath = prefix ? `${prefix}/${entry.name}` : entry.name;
      if (entry.isDirectory()) {
        files.push(...(await visit(resolve(path, entry.name), relativePath)));
      } else if (entry.isFile()) {
        fileCount++;
        if (fileCount > ATTACHMENT_MAX_ENTRIES)
          throw new AgentDocxError(
            "DOCX_TOO_LARGE",
            `Attachment bundle exceeds ${ATTACHMENT_MAX_ENTRIES} entries`,
          );
        const stat = await lstat(resolve(path, entry.name));
        totalBytes += stat.size;
        if (totalBytes > ATTACHMENT_MAX_TOTAL_BYTES)
          throw new AgentDocxError(
            "DOCX_TOO_LARGE",
            `Attachment bundle exceeds ${ATTACHMENT_MAX_TOTAL_BYTES / (1024 * 1024)} MiB decompressed`,
          );
        files.push(relativePath);
      } else {
        unsupported("Attachment bundle contains a non-regular entry");
      }
    }
    return files;
  };
  return visit(directory, "");
};

export const resolveAttachmentBundle = async (
  expected: AttachmentManifest | null,
  bundle: ImportAttachmentBundle | undefined,
): Promise<AttachmentResolution> => {
  if (!expected) {
    if (bundle) unsupported("DOCX has no external attachment inventory");
    return { assets: {}, inventory: {}, complete: true };
  }
  const expectedEntries = attachmentEntries(
    expected,
    "DOCX semantic attachment inventory",
  );
  const inventory = Object.fromEntries(
    expectedEntries.map((entry) => [
      entry.name,
      {
        sha256: entry.sha256,
        mediaType: entry.mediaType,
        bytes: entry.byteLength,
      },
    ]),
  );
  if (!bundle) return { assets: {}, inventory, complete: false };
  let supplied: AttachmentManifest;
  let sourceFiles: Readonly<Record<string, ImportedAsset>>;
  if ("directory" in bundle) {
    const manifestPath = pathInside(
      bundle.directory,
      "manifest.json",
      "Attachment manifest",
    );
    const content = await regularAttachmentFile(
      manifestPath,
      "Attachment manifest",
    );
    let parsed: unknown;
    try {
      parsed = JSON.parse(decodeDocxXml(content));
    } catch {
      unsupported("Attachment manifest is not valid JSON");
    }
    if (!validAttachmentManifest(parsed))
      unsupported("Attachment manifest does not have the version-1 shape");
    supplied = parsed as AttachmentManifest;
    const suppliedEntries = attachmentEntries(
      supplied,
      "Attachment manifest",
    );
    if (!sameAttachmentEntries(expectedEntries, suppliedEntries))
      unsupported(
        "Attachment bundle manifest does not exactly match the DOCX inventory",
      );
    const allowed = new Set([
      "manifest.json",
      ...suppliedEntries.map((entry) => entry.payloadPath),
    ]);
    const actual = await bundleFiles(bundle.directory);
    if (
      actual.length !== allowed.size ||
      actual.some((path) => !allowed.has(path))
    )
      unsupported("Attachment bundle has missing or extra payload files");
    const files: Record<string, ImportedAsset> = {};
    for (const entry of suppliedEntries) {
      const bytes = await regularAttachmentFile(
        pathInside(bundle.directory, entry.payloadPath, "Attachment payload"),
        `Attachment payload ${entry.name}`,
      );
      files[entry.name] = { bytes, mediaType: entry.mediaType };
    }
    sourceFiles = files;
  } else {
    supplied = bundle.manifest;
    const suppliedEntries = attachmentEntries(
      supplied,
      "Attachment manifest",
    );
    if (!sameAttachmentEntries(expectedEntries, suppliedEntries))
      unsupported(
        "Attachment bundle manifest does not exactly match the DOCX inventory",
      );
    if (
      Object.keys(bundle.files).length !== suppliedEntries.length ||
      suppliedEntries.some((entry) => bundle.files[entry.name] === undefined)
    )
      unsupported("Attachment bundle has missing or extra payload files");
    sourceFiles = bundle.files;
  }
  const assets: Record<string, ImportedAsset> = {};
  for (const entry of expectedEntries) {
    const asset = sourceFiles[entry.name];
    if (
      !asset ||
      typeof asset.mediaType !== "string" ||
      !(asset.bytes instanceof Uint8Array) ||
      asset.mediaType !== entry.mediaType ||
      asset.bytes.byteLength !== entry.byteLength ||
      sha256Hex(asset.bytes) !== entry.sha256
    )
      unsupported(
        `Attachment payload does not match manifest entry: ${entry.name}`,
      );
    assets[entry.name] = asset as ImportedAsset;
  }
  return { assets, inventory, complete: true };
};

const embeddedAssets = (
  parts: ReadonlyMap<string, Uint8Array>,
): readonly ImportedAsset[] =>
  [...parts.entries()]
    .filter(([path]) => path.startsWith("word/media/"))
    .map(([path, bytes]) => ({
      bytes,
      mediaType: path.endsWith(".png")
        ? "image/png"
        : /\.jpe?g$/i.test(path)
          ? "image/jpeg"
          : "application/octet-stream",
    }));

export const sourceAssetsForSemanticDocument = (
  semantic: SemanticManifest | null,
  external: AttachmentResolution,
  parts: ReadonlyMap<string, Uint8Array>,
): {
  assets: Readonly<Record<string, ImportedAsset>>;
  unresolved: readonly string[];
} => {
  if (!semantic) return { assets: {}, unresolved: [] };
  const documentAssets = asObject(
    semantic.document.assets ?? {},
    "Semantic manifest document assets",
  );
  const candidates = [
    ...embeddedAssets(parts),
    ...Object.values(external.assets),
  ];
  const assets: Record<string, ImportedAsset> = {};
  const unresolved: string[] = [];
  for (const [name, raw] of Object.entries(documentAssets)) {
    const asset = asObject(raw, `Semantic manifest asset ${name}`);
    exactKeys(
      asset,
      ["sha256", "mediaType", "bytes"],
      `Semantic manifest asset ${name}`,
    );
    if (
      typeof asset.sha256 !== "string" ||
      !/^sha256:[a-f0-9]{64}$/.test(asset.sha256) ||
      typeof asset.mediaType !== "string" ||
      !Number.isSafeInteger(asset.bytes) ||
      (asset.bytes as number) < 0
    )
      unsupported(`Semantic manifest asset ${name} is invalid`);
    const match = candidates.find(
      (candidate) =>
        candidate.mediaType === asset.mediaType &&
        candidate.bytes.byteLength === asset.bytes &&
        sha256Hex(candidate.bytes) === asset.sha256,
    );
    if (match) {
      assets[name] = match;
    } else {
      unresolved.push(name);
      assets[name] = {
        bytes: new Uint8Array(),
        mediaType: asset.mediaType as string,
      };
    }
  }
  return { assets, unresolved };
};
export const validAttachmentManifest = (
  value: unknown,
): value is AttachmentManifest => {
  if (value === null || typeof value !== "object" || Array.isArray(value))
    return false;
  const manifest = value as Record<string, unknown>;
  if (
    manifest.schemaVersion !== 1 ||
    !Array.isArray(manifest.entries) ||
    Object.keys(manifest).length !== 2 ||
    Object.keys(manifest).some(
      (key) => !["schemaVersion", "entries"].includes(key),
    )
  )
    return false;
  const names = new Set<string>();
  const paths = new Set<string>();
  return manifest.entries.every((entry) => {
    if (entry === null || typeof entry !== "object" || Array.isArray(entry))
      return false;
    const record = entry as Record<string, unknown>;
    if (
      Object.keys(record).length !== 5 ||
      !["name", "mediaType", "byteLength", "sha256", "payloadPath"].every(
        (key) => key in record,
      ) ||
      typeof record.name !== "string" ||
      typeof record.mediaType !== "string" ||
      !Number.isSafeInteger(record.byteLength) ||
      (record.byteLength as number) < 0 ||
      typeof record.sha256 !== "string" ||
      !/^sha256:[a-f0-9]{64}$/.test(record.sha256) ||
      typeof record.payloadPath !== "string"
    )
      return false;
    const name = record.name as string;
    const payloadPath = record.payloadPath as string;
    if (
      name.length === 0 ||
      name.startsWith("/") ||
      name.includes("\\") ||
      name.split("/").some((part) => part === "" || part === "." || part === "..") ||
      !payloadPath.startsWith("files/") ||
      payloadPath.includes("\\") ||
      payloadPath
        .split("/")
        .some(
          (part, index) =>
            index !== 0 && (part === "" || part === "." || part === ".."),
        ) ||
      names.has(name) ||
      paths.has(payloadPath)
    )
      return false;
    names.add(name);
    paths.add(payloadPath);
    return true;
  });
};

export const attachmentEntries = (
  manifest: AttachmentManifest,
  label: string,
): readonly AttachmentManifest["entries"][number][] => {
  if (!validAttachmentManifest(manifest))
    unsupported(
      `${label} does not have the version-1 attachment manifest shape`,
    );
  return [...manifest.entries].sort(
    (left, right) =>
      left.name.localeCompare(right.name) ||
      left.payloadPath.localeCompare(right.payloadPath),
  );
};