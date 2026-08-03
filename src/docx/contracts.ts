import type {
  MeasurementResult,
  PageCountSource,
  RendererMode,
  SourcePosition,
} from "../types.js";
import type {
  Actor,
  BlockId,
  FootnoteDefinition,
  LegalBlock,
  ReviewAnnotation,
  RevisionId,
} from "../legal/model.js";
import type { ValidationResult } from "../legal/rules.js";
import type { ProjectMeasurementResult } from "../project/contracts.js";
import type { ChangeSet, RevisionRecord } from "../revisions/types.js";

export type BodyBlockManifestEntry = {
  id: BlockId;
  bookmark: string;
  index: number;
  parentId: BlockId | null;
  depth: number;
  kind: LegalBlock["kind"];
  position: SourcePosition;
  preview: string;
};

export type RendererProvenance = {
  generator: "agent-docx";
  requested: RendererMode;
  pageCountSource: PageCountSource;
  wordVersion?: string;
  libreOfficeVersion?: string;
  verification?: {
    revisionCount: number;
    commentCount: number;
    fieldCount: number;
  };
};

export type AttachmentInventoryEntry = {
  name: string;
  mediaType: string;
  byteLength: number;
  sha256: RevisionId;
  payloadPath: string;
};

export type AttachmentManifest = {
  schemaVersion: 1;
  entries: readonly AttachmentInventoryEntry[];
};

export type GeneratedAttachmentBundle = {
  manifestSha256: RevisionId;
  manifest: AttachmentManifest;
  files: Readonly<Record<string, { bytes: Uint8Array; mediaType: string }>>;
};

export type ArtifactAttachmentBundle =
  | {
      path: null;
      storePath: null;
      manifestSha256: RevisionId;
      manifest: AttachmentManifest;
    }
  | {
      path: string;
      storePath: string;
      manifestSha256: RevisionId;
      manifest: AttachmentManifest;
    };

export type ArtifactResultBase = {
  schemaVersion: 1;
  mediaType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
  byteLength: number;
  sha256: RevisionId;
  provenanceSha256: RevisionId;
  documentId: string;
  profile: string;
  rulePack: string | null;
  rendererProvenance: RendererProvenance;
};

export type ArtifactResult =
  | (ArtifactResultBase & {
      path: null;
      storePath: null;
      attachments: Extract<ArtifactAttachmentBundle, { path: null }> | null;
      revision: null;
      mode: "clean";
      baseRevision: null;
    })
  | (ArtifactResultBase & {
      path: string;
      storePath: string;
      attachments: Extract<ArtifactAttachmentBundle, { path: string }> | null;
      revision: RevisionId;
      mode: "clean";
      baseRevision: null;
    })
  | (ArtifactResultBase & {
      path: string;
      storePath: string;
      attachments: Extract<ArtifactAttachmentBundle, { path: string }> | null;
      revision: RevisionId;
      mode: "redline";
      baseRevision: RevisionId;
    });

export type CompiledDocxBase = {
  schemaVersion: 1;
  bytes: Uint8Array;
  attachments: GeneratedAttachmentBundle | null;
  validation: ValidationResult;
  blocks: readonly BodyBlockManifestEntry[];
};

export type CompiledDocx =
  | (CompiledDocxBase & {
      measurement: Omit<MeasurementResult, "generatedDocx">;
      artifact: Extract<ArtifactResult, { path: null }>;
    })
  | (CompiledDocxBase & {
      measurement: ProjectMeasurementResult;
      artifact: Extract<ArtifactResult, { path: string }>;
    });

export type StatelessCompiledDocx = Extract<
  CompiledDocx,
  { artifact: { path: null } }
>;
export type ProjectCompiledDocx = Extract<
  CompiledDocx,
  { artifact: { path: string } }
>;
export type SerializableCompiledDocx = CompiledDocx extends infer Value
  ? Value extends unknown
    ? Omit<Value, "bytes" | "attachments">
    : never
  : never;

export type ExportRendererOptions = {
  renderer?: RendererMode;
  officeTimeoutMs?: number;
  word?: { powerShellPath?: string };
  libreoffice?: {
    executablePath?: string;
    installedFonts?: readonly { family: string; path: string }[];
  };
};

export type ExportDocxInput =
  | {
      revision: RevisionId | "HEAD";
      mode: "clean";
      output: string;
      options?: ExportRendererOptions;
    }
  | {
      revision: RevisionId | "HEAD";
      mode: "redline";
      baseRevision: RevisionId | "HEAD";
      output: string;
      options?: ExportRendererOptions;
    };

export type ImportAttachmentBundle =
  | { directory: string }
  | {
      files: Readonly<Record<string, { bytes: Uint8Array; mediaType: string }>>;
      manifest: AttachmentManifest;
    };

export type ImportDocxInput =
  | {
      input: string | Uint8Array;
      inspectOnly: true;
      attachments?: ImportAttachmentBundle;
    }
  | {
      input: string | Uint8Array;
      inspectOnly: false;
      attachments?: ImportAttachmentBundle;
      documentId: string;
      output: string;
      author: Actor;
      message: string;
    };

export type DocxFidelityItem<
  Status extends "preserved" | "normalized" | "externalized" | "unsupported",
> = {
  status: Status;
  partPath: string;
  relationshipId: string | null;
  ooxmlKind: string;
  count: number;
  blockIds: readonly BlockId[];
  sourcePositions: readonly SourcePosition[];
  explanation: string;
};

export type DocxImportBase<
  Status extends "preserved" | "normalized" | "externalized" | "unsupported",
> = {
  schemaVersion: 1;
  recognized: {
    blocks: readonly LegalBlock[];
    footnotes: readonly FootnoteDefinition[];
    annotations: readonly ReviewAnnotation[];
    assets: Readonly<
      Record<string, { sha256: RevisionId; mediaType: string; bytes: number }>
    >;
  };
  fidelity: {
    overall: Status extends "unsupported"
      ? "preserved" | "normalized" | "unsupported"
      : "preserved" | "normalized";
    items: readonly DocxFidelityItem<Status>[];
  };
};

export type DocxImportResult =
  | (DocxImportBase<
      "preserved" | "normalized" | "externalized" | "unsupported"
    > & {
      inspectOnly: true;
      mode: "inspect";
      output: null;
      sourceSha256: RevisionId | null;
      baseRevision: null;
      headRevision: null;
      revisions: readonly [];
    })
  | (DocxImportBase<"preserved" | "normalized" | "externalized"> & {
      inspectOnly: false;
      mode: "clean";
      output: string;
      sourceSha256: RevisionId;
      baseRevision: RevisionId;
      headRevision: RevisionId;
      revisions: readonly [RevisionId];
    })
  | (DocxImportBase<"preserved" | "normalized" | "externalized"> & {
      inspectOnly: false;
      mode: "tracked";
      output: string;
      sourceSha256: RevisionId;
      baseRevision: RevisionId;
      headRevision: RevisionId;
      revisions: readonly [RevisionId, RevisionId];
    });

export type GeneratedDocxOptions = {
  revision?: RevisionRecord;
  changeSet?: ChangeSet;
  annotations: readonly ReviewAnnotation[];
  validation: ValidationResult;
  dependencies: ReadonlyMap<
    string,
    { sha256: RevisionId; mediaType: string; bytes: Uint8Array }
  >;
};
