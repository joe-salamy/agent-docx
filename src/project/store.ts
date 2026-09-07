import type { RevisionId } from "../legal/model.js";
import type { AgentDocxManifest, DependencyHashes } from "./contracts.js";

export const STORE_DIR = ".agent-docx";
export const INIT_INTENT = ".agent-docx.init.json";
export const EXPORT_INTENT = ".agent-docx.export.json";
export const LOCK_NAME = ".agent-docx.lock";
export const BINDING_FILE = "project.json";

export type ExportIntent = {
  schemaVersion: 1;
  state: "preparing" | "prepared";
  projectId: string;
  manifestPath: string;
  owner: string;
  outputPath: string;
  attachmentPath: string | null;
  stagePath: string;
  docxStagePath: string;
  attachmentStagePath: string | null;
  pdfStagePath: string | null;
  docxSha256: RevisionId;
  attachmentManifestSha256: RevisionId | null;
  artifactProvenanceSha256: RevisionId;
  artifactStorePath: string;
  attachmentStorePath: string | null;
  pdfStorePath: string | null;
  pdfSha256: RevisionId | null;
};
export type OpenedStore = {
  manifestPath: string;
  projectDirectory: string;
  storePath: string;
  manifest: AgentDocxManifest;
};

export type ProjectSnapshot = {
  source: string;
  sourceObject: RevisionId;
  documentConfigObject: RevisionId;
  dependencyObjects: DependencyHashes;
  dependencyBytes: ReadonlyMap<
    string,
    { bytes: Uint8Array; mediaType: string }
  >;
  workingTreeHash: RevisionId;
};

export * from "./store/validate.js";
export * from "./store/objects.js";
export * from "./store/recovery.js";
export * from "./store/lock.js";
