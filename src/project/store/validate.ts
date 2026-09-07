import { AgentDocxError } from "../../types.js";
import { objectRecord } from "../../json-contract.js";
import {
  assertChrome as assertAgentChrome,
  assertMetadata as assertAgentMetadata,
} from "../../agent-protocol.js";
import { builtInProfiles } from "../../profiles.js";
import { isDocumentId } from "../../legal/model.js";
import { pathsOverlap } from "../fs-util.js";
import { assertRelativeManifestPath, assertWithin } from "./objects.js";
import { STORE_DIR } from "../store.js";
import type {
  AgentDocxDocumentConfig,
  AgentDocxManifest,
  FilingSet,
} from "../contracts.js";

const allowedManifestKeys: Record<string, true> = {
  schemaVersion: true,
  projectId: true,
  defaultDocument: true,
  storeDir: true,
  documents: true,
  filingSets: true,
};

const allowedDocumentKeys: Record<string, true> = {
  id: true,
  source: true,
  profile: true,
  filingKind: true,
  rulePack: true,
  rulePacks: true,
  template: true,
  assetsDir: true,
  fontSet: true,
  chrome: true,
  metadata: true,
};

const allowedFilingSetKeys: Record<string, true> = {
  id: true,
  label: true,
  documentIds: true,
  pageCap: true,
};

const hasOnlyKeys = (
  value: Record<string, unknown>,
  allowed: Record<string, true>,
): boolean => Object.keys(value).every((key) => allowed[key] === true);

const assertClosedKeys = (
  value: Record<string, unknown>,
  keys: readonly string[],
  name: string,
): void => {
  if (!Object.keys(value).every((key) => keys.includes(key)))
    throw new AgentDocxError(
      "PROJECT_INVALID",
      `${name} has unknown properties`,
    );
};

const assertUniqueMetadataIds = (value: unknown): void => {
  const metadata = value as Record<string, unknown>;
  for (const key of ["parties", "counsel"] as const) {
    const entries = metadata[key];
    if (!Array.isArray(entries)) continue;
    const ids = new Set<string>();
    for (const entry of entries) {
      const id = (entry as Record<string, unknown>).id;
      if (typeof id !== "string") continue;
      if (ids.has(id))
        throw new AgentDocxError(
          "PROJECT_INVALID",
          `Document metadata ${key} entry id is invalid`,
        );
      ids.add(id);
    }
  }
};

const assertMetadata = (value: unknown): void => {
  try {
    assertAgentMetadata(value);
  } catch (error) {
    if (error instanceof AgentDocxError && error.code === "INVALID_ARGUMENT")
      throw new AgentDocxError("PROJECT_INVALID", error.message);
    throw error;
  }
  assertUniqueMetadataIds(value);
};

const assertChrome = (value: unknown): void => {
  try {
    assertAgentChrome(value);
  } catch (error) {
    if (error instanceof AgentDocxError && error.code === "INVALID_ARGUMENT")
      throw new AgentDocxError("PROJECT_INVALID", error.message);
    throw error;
  }
};

const validateDocumentConfig = (value: unknown): AgentDocxDocumentConfig => {
  if (!value || typeof value !== "object" || Array.isArray(value))
    throw new AgentDocxError(
      "PROJECT_INVALID",
      "Document config must be an object",
    );
  const config = value as Record<string, unknown>;
  if (!hasOnlyKeys(config, allowedDocumentKeys))
    throw new AgentDocxError(
      "PROJECT_INVALID",
      "Document config has unknown properties",
    );
  if (typeof config.id !== "string" || !isDocumentId(config.id))
    throw new AgentDocxError("PROJECT_INVALID", "Document id is invalid");
  if (typeof config.source !== "string")
    throw new AgentDocxError("PROJECT_INVALID", "Document source is invalid");
  assertRelativeManifestPath(config.source, "Document source");
  if (
    typeof config.profile !== "string" ||
    !(config.profile in builtInProfiles)
  )
    throw new AgentDocxError("PROJECT_INVALID", "Document profile is invalid");
  assertMetadata(config.metadata);
  if (
    config.filingKind !== undefined &&
    ![
      "principal-brief",
      "reply-brief",
      "motion-document",
      "opposition-text",
      "reply-text",
    ].includes(String(config.filingKind))
  )
    throw new AgentDocxError(
      "PROJECT_INVALID",
      "Document filingKind is invalid",
    );
  if (
    config.rulePack !== undefined &&
    !["frap-32@2024-12-01", "cand-civil@2026-05-01"].includes(
      String(config.rulePack),
    )
  )
    throw new AgentDocxError("PROJECT_INVALID", "Document rulePack is invalid");
  if (config.rulePacks !== undefined) {
    if (!Array.isArray(config.rulePacks))
      throw new AgentDocxError(
        "PROJECT_INVALID",
        "Document rulePacks is invalid",
      );
    for (const [index, configuredPath] of config.rulePacks.entries()) {
      if (
        typeof configuredPath !== "string" ||
        configuredPath.length === 0 ||
        configuredPath.includes("\0")
      )
        throw new AgentDocxError(
          "PATH_OUTSIDE_PROJECT",
          `Document rulePacks[${index}] is invalid`,
        );
      assertRelativeManifestPath(
        configuredPath,
        `Document rulePacks[${index}]`,
      );
    }
  }
  if (config.chrome !== undefined) assertChrome(config.chrome);
  for (const name of ["template", "assetsDir"] as const) {
    const path = config[name];
    if (path !== undefined) {
      if (typeof path !== "string")
        throw new AgentDocxError("PROJECT_INVALID", `${name} is invalid`);
      assertRelativeManifestPath(path, name);
    }
  }
  if (config.fontSet !== undefined) {
    if (
      !config.fontSet ||
      typeof config.fontSet !== "object" ||
      Array.isArray(config.fontSet)
    )
      throw new AgentDocxError("PROJECT_INVALID", "fontSet is invalid");
    const fontSet = objectRecord(config.fontSet, "fontSet", {
      code: "PROJECT_INVALID",
    });
    assertClosedKeys(
      fontSet,
      ["family", "regularPath", "boldPath", "italicPath", "boldItalicPath"],
      "fontSet",
    );
    if (typeof fontSet.family !== "string" || fontSet.family.length === 0)
      throw new AgentDocxError("PROJECT_INVALID", "fontSet family is invalid");
    if (typeof fontSet.regularPath !== "string")
      throw new AgentDocxError(
        "PROJECT_INVALID",
        "fontSet regularPath is invalid",
      );
    for (const role of [
      "regularPath",
      "boldPath",
      "italicPath",
      "boldItalicPath",
    ] as const) {
      const path = fontSet[role];
      if (path !== undefined) {
        if (typeof path !== "string")
          throw new AgentDocxError(
            "PROJECT_INVALID",
            `fontSet ${role} is invalid`,
          );
        assertRelativeManifestPath(path, `fontSet ${role}`);
      }
    }
  }
  return config as unknown as AgentDocxDocumentConfig;
};

const validateManifestFilingSet = (
  value: unknown,
  documentIds: ReadonlySet<string>,
): FilingSet => {
  const set = objectRecord(value, "Filing set", { code: "PROJECT_INVALID" });
  if (!hasOnlyKeys(set, allowedFilingSetKeys))
    throw new AgentDocxError(
      "PROJECT_INVALID",
      "Filing set has unknown properties",
    );
  const id = set.id;
  if (typeof id !== "string" || !isDocumentId(id))
    throw new AgentDocxError("PROJECT_INVALID", "Filing set id is invalid");
  if (
    set.label !== undefined &&
    (typeof set.label !== "string" || set.label.length === 0)
  )
    throw new AgentDocxError(
      "PROJECT_INVALID",
      `Filing set ${id} label is invalid`,
    );
  if (!Array.isArray(set.documentIds) || set.documentIds.length === 0)
    throw new AgentDocxError(
      "PROJECT_INVALID",
      `Filing set ${id} documentIds must be a nonempty array`,
    );
  const references: string[] = [];
  const seen = new Set<string>();
  for (const value of set.documentIds) {
    if (typeof value !== "string")
      throw new AgentDocxError(
        "PROJECT_INVALID",
        `Filing set ${id} document id is invalid`,
      );
    if (seen.has(value))
      throw new AgentDocxError(
        "PROJECT_INVALID",
        `Filing set ${id} references duplicate document ${value}`,
      );
    if (!documentIds.has(value))
      throw new AgentDocxError(
        "PROJECT_INVALID",
        `Filing set ${id} references unknown document ${value}`,
      );
    seen.add(value);
    references.push(value);
  }
  if (
    set.pageCap !== undefined &&
    (!Number.isInteger(set.pageCap) || (set.pageCap as number) < 1)
  )
    throw new AgentDocxError(
      "PROJECT_INVALID",
      `Filing set ${id} pageCap is invalid`,
    );
  const label = set.label as string | undefined;
  const pageCap = set.pageCap as number | undefined;
  return {
    id,
    ...(label !== undefined ? { label } : {}),
    documentIds: references,
    ...(pageCap !== undefined ? { pageCap } : {}),
  };
};

export const validateManifest = (value: unknown): AgentDocxManifest => {
  if (!value || typeof value !== "object" || Array.isArray(value))
    throw new AgentDocxError(
      "PROJECT_INVALID",
      "Project manifest must be an object",
    );
  const manifest = value as Record<string, unknown>;
  if (!hasOnlyKeys(manifest, allowedManifestKeys))
    throw new AgentDocxError(
      "PROJECT_INVALID",
      "Project manifest has unknown properties",
    );
  if (manifest.schemaVersion !== 1)
    throw new AgentDocxError(
      "PROJECT_INVALID",
      "Project manifest schemaVersion must be 1",
    );
  if (
    typeof manifest.projectId !== "string" ||
    !/^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/.test(
      manifest.projectId,
    )
  )
    throw new AgentDocxError("PROJECT_INVALID", "Project id is invalid");
  if (
    typeof manifest.defaultDocument !== "string" ||
    !isDocumentId(manifest.defaultDocument)
  )
    throw new AgentDocxError(
      "PROJECT_INVALID",
      "Default document id is invalid",
    );
  if (manifest.storeDir !== STORE_DIR)
    throw new AgentDocxError(
      "PROJECT_INVALID",
      `Project storeDir must be ${STORE_DIR}`,
    );
  if (!Array.isArray(manifest.documents) || manifest.documents.length === 0)
    throw new AgentDocxError(
      "PROJECT_INVALID",
      "Project must contain at least one document",
    );
  const documents = manifest.documents.map(validateDocumentConfig);
  const ids = new Set<string>();
  const paths = new Set<string>();
  for (const document of documents) {
    if (ids.has(document.id))
      throw new AgentDocxError(
        "PROJECT_INVALID",
        `Duplicate document id: ${document.id}`,
      );
    if (paths.has(document.source))
      throw new AgentDocxError(
        "PROJECT_INVALID",
        `Duplicate document source: ${document.source}`,
      );
    ids.add(document.id);
    paths.add(document.source);
  }
  if (!ids.has(manifest.defaultDocument))
    throw new AgentDocxError(
      "PROJECT_INVALID",
      "Default document is not declared",
    );
  const filingSets =
    manifest.filingSets === undefined
      ? []
      : Array.isArray(manifest.filingSets)
        ? manifest.filingSets.map((entry) =>
            validateManifestFilingSet(entry, ids),
          )
        : (() => {
            throw new AgentDocxError(
              "PROJECT_INVALID",
              "Project filingSets must be an array",
            );
          })();
  const filingSetIds = new Set<string>();
  for (const filingSet of filingSets) {
    if (filingSetIds.has(filingSet.id))
      throw new AgentDocxError(
        "PROJECT_INVALID",
        `Duplicate filing set id: ${filingSet.id}`,
      );
    filingSetIds.add(filingSet.id);
  }
  return {
    schemaVersion: 1,
    projectId: manifest.projectId,
    defaultDocument: manifest.defaultDocument,
    storeDir: STORE_DIR,
    documents,
    ...(filingSets.length > 0 ? { filingSets } : {}),
  };
};
type OwnedManifestPath = {
  path: string;
  documentId: string;
  shareableFont: boolean;
};

export const validateManifestPaths = async (
  projectDirectory: string,
  manifest: AgentDocxManifest,
): Promise<void> => {
  const owned: OwnedManifestPath[] = [];
  const add = (
    path: string | undefined,
    label: string,
    documentId: string,
    shareableFont = false,
  ): void => {
    if (path === undefined) return;
    const normalized = assertRelativeManifestPath(path, label);
    const absolute = assertWithin(projectDirectory, normalized, label);
    for (const existing of owned) {
      const sharedFont =
        absolute === existing.path && shareableFont && existing.shareableFont;
      if (pathsOverlap(absolute, existing.path) && !sharedFont)
        throw new AgentDocxError(
          "PROJECT_INVALID",
          `Manifest path is already owned: ${path}`,
        );
    }
    owned.push({ path: absolute, documentId, shareableFont });
  };
  for (const document of manifest.documents) {
    add(document.source, "Document source", document.id);
    add(document.template, "Document template", document.id);
    add(document.assetsDir, "Assets directory", document.id);
    for (const [index, path] of (document.rulePacks ?? []).entries())
      add(path, `Rule pack ${index}`, document.id);
    for (const [role, path] of [
      ["regular", document.fontSet?.regularPath],
      ["bold", document.fontSet?.boldPath],
      ["italic", document.fontSet?.italicPath],
      ["boldItalic", document.fontSet?.boldItalicPath],
    ] as const)
      add(path, `Font ${role}`, document.id, true);
  }
};
