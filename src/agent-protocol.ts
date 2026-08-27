import { dirname, isAbsolute, relative, resolve, sep } from "node:path";
import { publicPath } from "./path-util.js";
import type {
  DocumentConfigUpdate,
  ProjectDocumentInput,
  ResolveChangesInput,
} from "./project/contracts.js";
import type { Actor, RevisionId } from "./legal/model.js";
import type { ChangeSet } from "./revisions/types.js";
import type { SourcePatch } from "./draft/types.js";
import { AgentDocxError } from "./types.js";
import { objectRecord } from "./json-contract.js";

export const agentActions = [
  "project.init",
  "project.add",
  "project.get",
  "document.configure",
  "document.get",
  "document.measure",
  "document.validate",
  "revision.checkpoint",
  "revision.list",
  "revision.get",
  "revision.restore",
  "revision.diff",
  "revision.resolve",
  "draft.guidance",
  "draft.evaluate",
  "draft.apply",
  "review.add",
  "review.resolve",
  "docx.export",
  "docx.import",
  "docx.inspect",
  "docx.importRedline",
  "filingSet.add",
  "filingSet.remove",
  "filingSet.get",
  "filingSet.validate",
] as const;

export type AgentAction = (typeof agentActions)[number];
export type AgentRequestId = string | number | null;
export type AgentParams = Record<string, unknown>;

export type AgentFontSet = AgentParams & {
  family: string;
  regularPath: string;
  boldPath?: string;
  italicPath?: string;
  boldItalicPath?: string;
};
export type AgentRange = AgentParams & { start: number; length: number };
export type AgentConfigUpdate = AgentParams & DocumentConfigUpdate;
export type AgentRendererOptions = AgentParams & {
  renderer?: "deterministic" | "word" | "libreoffice" | "compare";
  officeTimeoutMs?: number;
  paragraphDiagnostics?: boolean;
  sectionDiagnostics?: boolean;
  lineDiagnostics?: boolean;
  trim?: false | AgentParams;
  word?: AgentParams;
  libreoffice?: AgentParams;
};

export type ProjectInitParams = AgentParams & {
  documentId: string;
  source: string;
  createSource?: boolean;
  profile: string;
  filingKind?: string;
  rulePack?: string;
  rulePacks?: string[];
  template?: string;
  assetsDir?: string;
  fontSet?: AgentFontSet;
  metadata: AgentParams;
  chrome?: AgentParams;
};
export type ProjectAddParams = ProjectInitParams & { makeDefault?: boolean };
export type NoOptionsParams = AgentParams;
export type DocumentConfigureParams = AgentParams & {
  documentId: string;
  baseRevision: RevisionId | "HEAD" | null;
  changes: AgentConfigUpdate;
  author: Actor;
  message: string;
};
export type DocumentParams = AgentParams & {
  documentId: string;
  revision?: RevisionId | "HEAD";
};
export type DocumentMeasureParams = DocumentParams & {
  options?: AgentRendererOptions;
};
export type RevisionCheckpointParams = AgentParams & {
  documentId: string;
  baseRevision: RevisionId | "HEAD" | null;
  author: Actor;
  message: string;
};
export type RevisionListParams = AgentParams & {
  documentId: string;
  limit?: number;
  cursor?: RevisionId;
};
export type RevisionGetParams = AgentParams & {
  documentId: string;
  revision: RevisionId | "HEAD";
};
export type RevisionRestoreParams = AgentParams & {
  documentId: string;
  baseRevision: RevisionId | "HEAD";
  targetRevision: RevisionId | "HEAD";
  author: Actor;
  message: string;
};
export type RevisionDiffParams = AgentParams & {
  documentId: string;
  baseRevision: RevisionId | "HEAD";
  headRevision: RevisionId | "HEAD";
  output?: string;
};
export type RevisionResolveParams = AgentParams & {
  documentId: string;
  changeSet: ChangeSet;
  decisions: ResolveChangesInput["decisions"];
  author: Actor;
  message: string;
};
export type DraftEvaluateParams = AgentParams & {
  patch: SourcePatch;
  renderer?: "deterministic" | "word" | "libreoffice" | "compare";
};
export type DraftApplyParams = AgentParams & {
  patch: SourcePatch;
  patchHash: RevisionId;
  gate?: "report" | "not-worse" | "pass";
  author: Actor;
  message: string;
};
export type ReviewAddParams = AgentParams & {
  documentId: string;
  revision: RevisionId | "HEAD";
  blockId: string;
  range?: AgentRange;
  author: Actor;
  message: string;
};
export type ReviewResolveParams = AgentParams & {
  documentId: string;
  revision: RevisionId | "HEAD";
  annotationId: string;
  author: Actor;
  message: string;
};
export type DocxExportParams = AgentParams & {
  documentId: string;
  revision: RevisionId | "HEAD";
  mode: "clean" | "redline" | "pdf";
  baseRevision?: RevisionId | "HEAD";
  output: string;
  options?: AgentRendererOptions;
};
export type DocxImportParams = AgentParams & {
  input: string;
  attachments?: string;
  inspectOnly: boolean;
  documentId?: string;
  output?: string;
  author?: Actor;
  message?: string;
};
export type DocxInspectParams = AgentParams & { input: string };
export type DocxImportRedlineParams = AgentParams & {
  documentId: string;
  input: string;
  attachments?: string;
  author: Actor;
  message: string;
};
export type FilingSetAddParams = AgentParams & {
  id: string;
  label?: string;
  documentIds: string[];
  pageCap?: number;
};
export type FilingSetParams = AgentParams & { id: string };

type AgentRequestVariant<A extends AgentAction, P extends AgentParams> = {
  schemaVersion: 1;
  id?: AgentRequestId;
  action: A;
  project?: string;
  params: P;
};

export type AgentRequest =
  | AgentRequestVariant<"project.init", ProjectInitParams>
  | AgentRequestVariant<"project.add", ProjectAddParams>
  | AgentRequestVariant<"project.get", NoOptionsParams>
  | AgentRequestVariant<"document.configure", DocumentConfigureParams>
  | AgentRequestVariant<"document.get", DocumentParams>
  | AgentRequestVariant<"document.measure", DocumentMeasureParams>
  | AgentRequestVariant<"document.validate", DocumentParams>
  | AgentRequestVariant<"revision.checkpoint", RevisionCheckpointParams>
  | AgentRequestVariant<"revision.list", RevisionListParams>
  | AgentRequestVariant<"revision.get", RevisionGetParams>
  | AgentRequestVariant<"revision.restore", RevisionRestoreParams>
  | AgentRequestVariant<"revision.diff", RevisionDiffParams>
  | AgentRequestVariant<"revision.resolve", RevisionResolveParams>
  | AgentRequestVariant<"draft.guidance", DocumentParams>
  | AgentRequestVariant<"draft.evaluate", DraftEvaluateParams>
  | AgentRequestVariant<"draft.apply", DraftApplyParams>
  | AgentRequestVariant<"review.add", ReviewAddParams>
  | AgentRequestVariant<"review.resolve", ReviewResolveParams>
  | AgentRequestVariant<"docx.export", DocxExportParams>
  | AgentRequestVariant<"docx.import", DocxImportParams>
  | AgentRequestVariant<"docx.inspect", DocxInspectParams>
  | AgentRequestVariant<"docx.importRedline", DocxImportRedlineParams>
  | AgentRequestVariant<"filingSet.add", FilingSetAddParams>
  | AgentRequestVariant<"filingSet.remove", FilingSetParams>
  | AgentRequestVariant<"filingSet.get", FilingSetParams>
  | AgentRequestVariant<"filingSet.validate", FilingSetParams>;

export type AgentParamsForAction<A extends AgentAction> = Extract<
  AgentRequest,
  { action: A }
>["params"];

export type AgentDispatchResult = {
  request: AgentRequest;
  project: string | null;
  documentId: string | null;
  revision: RevisionId | null;
  value: unknown;
};

export const assertRecord = (
  value: unknown,
  label: string,
): Record<string, unknown> => objectRecord(value, label);

export const assertKeys = (
  value: Record<string, unknown>,
  allowed: readonly string[],
  label: string,
): void => {
  const unknown = Object.keys(value).find((key) => !allowed.includes(key));
  if (unknown)
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `${label} has unknown property: ${unknown}`,
    );
};

export const requiredString = (
  value: Record<string, unknown>,
  key: string,
): string => {
  if (typeof value[key] !== "string" || value[key] === "")
    throw new AgentDocxError("INVALID_ARGUMENT", `${key} is required`);
  return value[key] as string;
};

export const optionalString = (
  value: Record<string, unknown>,
  key: string,
): string | undefined => {
  if (value[key] === undefined) return undefined;
  if (typeof value[key] !== "string" || value[key] === "")
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `${key} must be a non-empty string`,
    );
  return value[key] as string;
};

export const optionalBoolean = (
  value: Record<string, unknown>,
  key: string,
): boolean | undefined => {
  if (value[key] === undefined) return undefined;
  if (typeof value[key] !== "boolean")
    throw new AgentDocxError("INVALID_ARGUMENT", `${key} must be a boolean`);
  return value[key] as boolean;
};

export const asRevision = (
  value: unknown,
  key: string,
  nullable = false,
): RevisionId | "HEAD" | null => {
  if (nullable && value === null) return null;
  if (
    value === "HEAD" ||
    (typeof value === "string" && /^sha256:[0-9a-f]{64}$/.test(value))
  )
    return value as RevisionId | "HEAD";
  throw new AgentDocxError(
    "INVALID_ARGUMENT",
    `${key} must be HEAD or a sha256 revision ID`,
  );
};

export const actor = (value: unknown): Actor => {
  const candidate = assertRecord(value, "author");
  assertKeys(candidate, ["name", "email"], "author");
  const name = requiredString(candidate, "name");
  const email = optionalString(candidate, "email");
  return email === undefined ? { name } : { name, email };
};

export const invocationPath = (cwd: string, path: string): string =>
  isAbsolute(path) ? path : resolve(cwd, path);

export { publicPath };

export const manifestRelativePath = (
  cwd: string,
  manifestPath: string,
  path: string,
): string => {
  const candidate = relative(dirname(manifestPath), invocationPath(cwd, path))
    .split(sep)
    .join("/");
  if (
    !candidate ||
    candidate === ".." ||
    candidate.startsWith("../") ||
    isAbsolute(candidate)
  )
    throw new AgentDocxError("PATH_OUTSIDE_PROJECT", "Path is outside project");
  return candidate;
};

export const projectPath = (
  cwd: string,
  requested: string | undefined,
): string => invocationPath(cwd, requested ?? "agent-docx.json");

export const projectInput = (
  cwd: string,
  manifestPath: string,
  params: Record<string, unknown>,
  includeDefault: boolean,
): ProjectDocumentInput & { makeDefault?: boolean } => {
  assertKeys(
    params,
    [
      "documentId",
      "source",
      "createSource",
      "profile",
      "filingKind",
      "rulePack",
      "rulePacks",
      "template",
      "assetsDir",
      "fontSet",
      "metadata",
      "chrome",
      ...(includeDefault ? ["makeDefault"] : []),
    ],
    "project document parameters",
  );
  const documentId = requiredDocumentId(params, "documentId");
  const source = manifestRelativePath(
    cwd,
    manifestPath,
    requiredString(params, "source"),
  );
  const profile = requiredEnum(params, "profile", profileIds);
  assertMetadata(params.metadata);
  const metadata = assertRecord(params.metadata, "metadata");
  const createSource = optionalBoolean(params, "createSource");
  const template = optionalString(params, "template");
  const assetsDir = optionalString(params, "assetsDir");
  const filingKind = optionalEnum(params, "filingKind", filingKinds);
  const rulePack = optionalEnum(params, "rulePack", rulePackIds);
  const rulePacks = optionalRulePacks(params);
  if (params.chrome !== undefined) assertChrome(params.chrome);
  const chrome =
    params.chrome === undefined
      ? undefined
      : assertRecord(params.chrome, "chrome");
  let fontSet: ProjectDocumentInput["fontSet"];
  if (params.fontSet !== undefined) {
    const raw = assertRecord(params.fontSet, "fontSet");
    assertFontSet(raw);
    assertKeys(
      raw,
      ["family", "regularPath", "boldPath", "italicPath", "boldItalicPath"],
      "fontSet",
    );
    const family = requiredString(raw, "family");
    const regularPath = manifestRelativePath(
      cwd,
      manifestPath,
      requiredString(raw, "regularPath"),
    );
    const boldPath = optionalString(raw, "boldPath");
    const italicPath = optionalString(raw, "italicPath");
    const boldItalicPath = optionalString(raw, "boldItalicPath");
    fontSet = {
      family,
      regularPath,
      ...(boldPath
        ? { boldPath: manifestRelativePath(cwd, manifestPath, boldPath) }
        : {}),
      ...(italicPath
        ? { italicPath: manifestRelativePath(cwd, manifestPath, italicPath) }
        : {}),
      ...(boldItalicPath
        ? {
            boldItalicPath: manifestRelativePath(
              cwd,
              manifestPath,
              boldItalicPath,
            ),
          }
        : {}),
    };
  }
  return {
    documentId,
    source,
    profile: profile as NonNullable<ProjectDocumentInput["profile"]>,
    metadata: metadata as NonNullable<ProjectDocumentInput["metadata"]>,
    ...(createSource === undefined ? {} : { createSource }),
    ...(template
      ? { template: manifestRelativePath(cwd, manifestPath, template) }
      : {}),
    ...(assetsDir
      ? { assetsDir: manifestRelativePath(cwd, manifestPath, assetsDir) }
      : {}),
    ...(filingKind
      ? {
          filingKind: filingKind as NonNullable<
            ProjectDocumentInput["filingKind"]
          >,
        }
      : {}),
    ...(rulePack
      ? { rulePack: rulePack as NonNullable<ProjectDocumentInput["rulePack"]> }
      : {}),
    ...(rulePacks
      ? {
          rulePacks: rulePacks.map((path) =>
            manifestRelativePath(cwd, manifestPath, path),
          ),
        }
      : {}),
    ...(fontSet ? { fontSet } : {}),
    ...(chrome
      ? { chrome: chrome as NonNullable<ProjectDocumentInput["chrome"]> }
      : {}),
    ...(includeDefault && optionalBoolean(params, "makeDefault")
      ? { makeDefault: true }
      : {}),
  };
};
export const configUpdate = (
  cwd: string,
  manifestPath: string,
  value: unknown,
): DocumentConfigUpdate => {
  assertConfigUpdate(value);
  const raw = assertRecord(value, "changes");
  const changes: DocumentConfigUpdate = {};
  if (raw.profile !== undefined)
    changes.profile = raw.profile as NonNullable<
      DocumentConfigUpdate["profile"]
    >;
  if (raw.filingKind !== undefined)
    changes.filingKind = raw.filingKind as NonNullable<
      DocumentConfigUpdate["filingKind"]
    >;
  if (raw.rulePack !== undefined)
    changes.rulePack = raw.rulePack as NonNullable<
      DocumentConfigUpdate["rulePack"]
    >;
  if (raw.rulePacks !== undefined)
    changes.rulePacks =
      raw.rulePacks === null
        ? null
        : optionalRulePacks({ rulePacks: raw.rulePacks })!.map((path) =>
            manifestRelativePath(cwd, manifestPath, path),
          );
  if (raw.template !== undefined)
    changes.template =
      raw.template === null
        ? null
        : manifestRelativePath(cwd, manifestPath, raw.template as string);
  if (raw.assetsDir !== undefined)
    changes.assetsDir =
      raw.assetsDir === null
        ? null
        : manifestRelativePath(cwd, manifestPath, raw.assetsDir as string);
  if (raw.fontSet !== undefined) {
    if (raw.fontSet === null) changes.fontSet = null;
    else {
      const fontSet = assertRecord(raw.fontSet, "changes.fontSet");
      changes.fontSet = {
        family: requiredString(fontSet, "family"),
        regularPath: manifestRelativePath(
          cwd,
          manifestPath,
          requiredString(fontSet, "regularPath"),
        ),
        ...(fontSet.boldPath === undefined
          ? {}
          : {
              boldPath: manifestRelativePath(
                cwd,
                manifestPath,
                requiredString(fontSet, "boldPath"),
              ),
            }),
        ...(fontSet.italicPath === undefined
          ? {}
          : {
              italicPath: manifestRelativePath(
                cwd,
                manifestPath,
                requiredString(fontSet, "italicPath"),
              ),
            }),
        ...(fontSet.boldItalicPath === undefined
          ? {}
          : {
              boldItalicPath: manifestRelativePath(
                cwd,
                manifestPath,
                requiredString(fontSet, "boldItalicPath"),
              ),
            }),
      };
    }
  }
  if (raw.metadata !== undefined)
    changes.metadata = raw.metadata as NonNullable<
      DocumentConfigUpdate["metadata"]
    >;
  if (raw.chrome !== undefined)
    changes.chrome = raw.chrome as NonNullable<DocumentConfigUpdate["chrome"]>;
  return changes;
};

export const asPatch = (value: unknown): SourcePatch => {
  assertSourcePatch(value);
  return value as SourcePatch;
};

export const asChangeSet = (value: unknown): ChangeSet => {
  assertChangeSet(value);
  return value as ChangeSet;
};

export const noOptions = (
  value: unknown,
  label: string,
): Record<string, never> => {
  const result = assertRecord(value, label);
  assertKeys(result, [], label);
  return result as Record<string, never>;
};
export const documentIdPattern = /^[a-z0-9][a-z0-9-]{0,63}$/;
export const blockIdPattern =
  /^b_[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/;
export const annotationIdPattern =
  /^a_[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/;
export const revisionIdPattern = /^sha256:[0-9a-f]{64}$/;
export const changeIdPattern = /^c_[0-9a-f]{64}$/;

export const profileIds = [
  "us-district-conventional",
  "frap-32",
  "cand-civil",
] as const;
export const filingKinds = [
  "principal-brief",
  "reply-brief",
  "motion-document",
  "opposition-text",
  "reply-text",
] as const;
export const rulePackIds = [
  "frap-32@2024-12-01",
  "cand-civil@2026-05-01",
] as const;
export const rendererModes = [
  "deterministic",
  "word",
  "libreoffice",
  "compare",
] as const;

const hasOwn = (value: Record<string, unknown>, key: string): boolean =>
  Object.prototype.hasOwnProperty.call(value, key);

export const requiredDocumentId = (
  value: Record<string, unknown>,
  key: string,
): string => {
  const documentId = requiredString(value, key);
  if (!documentIdPattern.test(documentId))
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `${key} must be a document ID`,
    );
  return documentId;
};

export const requiredEnum = <T extends readonly string[]>(
  value: Record<string, unknown>,
  key: string,
  allowed: T,
): T[number] => {
  const candidate = requiredString(value, key);
  if (!(allowed as readonly string[]).includes(candidate))
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `${key} must be one of: ${allowed.join(", ")}`,
    );
  return candidate as T[number];
};

export const optionalEnum = <T extends readonly string[]>(
  value: Record<string, unknown>,
  key: string,
  allowed: T,
): T[number] | undefined => {
  if (value[key] === undefined) return undefined;
  return requiredEnum(value, key, allowed);
};

export const requiredRevisionId = (
  value: Record<string, unknown>,
  key: string,
): RevisionId => {
  const revision = requiredString(value, key);
  if (!revisionIdPattern.test(revision))
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `${key} must be a sha256 revision ID`,
    );
  return revision as RevisionId;
};

export const optionalRevision = (
  value: Record<string, unknown>,
  key: string,
): RevisionId | "HEAD" | undefined =>
  value[key] === undefined
    ? undefined
    : (asRevision(value[key], key) as RevisionId | "HEAD");

export const optionalRulePacks = (
  value: Record<string, unknown>,
): string[] | undefined => {
  if (value.rulePacks === undefined) return undefined;
  if (
    !Array.isArray(value.rulePacks) ||
    value.rulePacks.length === 0 ||
    value.rulePacks.some(
      (entry) => typeof entry !== "string" || entry.length === 0,
    )
  )
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      "rulePacks must be a non-empty array of pack file paths",
    );
  return value.rulePacks as string[];
};

export const requiredInteger = (
  value: Record<string, unknown>,
  key: string,
  minimum = 0,
  maximum = Number.MAX_SAFE_INTEGER,
): number => {
  const candidate = value[key];
  if (
    !Number.isInteger(candidate) ||
    (candidate as number) < minimum ||
    (candidate as number) > maximum
  )
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `${key} must be an integer from ${minimum} through ${maximum}`,
    );
  return candidate as number;
};

export const optionalInteger = (
  value: Record<string, unknown>,
  key: string,
  minimum = 0,
  maximum = Number.MAX_SAFE_INTEGER,
): number | undefined =>
  value[key] === undefined
    ? undefined
    : requiredInteger(value, key, minimum, maximum);

export const requiredText = (
  value: Record<string, unknown>,
  key: string,
  label: string,
): string => {
  if (typeof value[key] !== "string")
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `${label}.${key} must be a string`,
    );
  return value[key] as string;
};

export const optionalText = (
  value: Record<string, unknown>,
  key: string,
  label: string,
): string | undefined => {
  if (value[key] === undefined) return undefined;
  return requiredText(value, key, label);
};

const assertDate = (value: string, label: string): void => {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value))
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `${label} must use YYYY-MM-DD`,
    );
  const parsed = new Date(`${value}T00:00:00.000Z`);
  if (
    Number.isNaN(parsed.valueOf()) ||
    parsed.toISOString().slice(0, 10) !== value
  )
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `${label} must be a valid date`,
    );
};

export const assertMetadata = (value: unknown): void => {
  const metadata = assertRecord(value, "metadata");
  assertKeys(
    metadata,
    [
      "court",
      "jurisdiction",
      "caseName",
      "docketNumber",
      "documentTitle",
      "filingDate",
      "parties",
      "counsel",
      "certificates",
    ],
    "metadata",
  );
  for (const key of [
    "court",
    "jurisdiction",
    "caseName",
    "docketNumber",
    "documentTitle",
  ] as const)
    requiredText(metadata, key, "metadata");
  if (metadata.filingDate !== undefined)
    assertDate(
      requiredText(metadata, "filingDate", "metadata"),
      "metadata.filingDate",
    );

  const parties = metadata.parties;
  if (!Array.isArray(parties))
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      "metadata.parties must be an array",
    );
  for (const party of parties) {
    const record = assertRecord(party, "metadata party");
    assertKeys(record, ["id", "name", "role"], "metadata party");
    requiredDocumentId(record, "id");
    requiredText(record, "name", "metadata party");
    requiredText(record, "role", "metadata party");
  }

  const counsel = metadata.counsel;
  if (!Array.isArray(counsel))
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      "metadata.counsel must be an array",
    );
  for (const entry of counsel) {
    const record = assertRecord(entry, "metadata counsel");
    assertKeys(
      record,
      ["id", "name", "barNumber", "firm", "addressLines", "phone", "email"],
      "metadata counsel",
    );
    requiredDocumentId(record, "id");
    requiredText(record, "name", "metadata counsel");
    for (const key of ["barNumber", "firm", "phone", "email"] as const)
      optionalText(record, key, "metadata counsel");
    if (record.addressLines !== undefined) {
      if (!Array.isArray(record.addressLines))
        throw new AgentDocxError(
          "INVALID_ARGUMENT",
          "metadata counsel addressLines must be an array",
        );
      for (const line of record.addressLines)
        if (typeof line !== "string")
          throw new AgentDocxError(
            "INVALID_ARGUMENT",
            "metadata counsel addressLines must contain strings",
          );
    }
  }

  const certificates = metadata.certificates;
  if (!Array.isArray(certificates))
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      "metadata.certificates must be an array",
    );
  for (const certificate of certificates) {
    const record = assertRecord(certificate, "metadata certificate");
    if (record.kind === "service") {
      assertKeys(
        record,
        [
          "id",
          "kind",
          "statement",
          "servedOn",
          "method",
          "date",
          "signerCounselId",
        ],
        "service certificate",
      );
      requiredDocumentId(record, "id");
      requiredText(record, "statement", "service certificate");
      requiredText(record, "method", "service certificate");
      requiredDocumentId(record, "signerCounselId");
      if (
        !Array.isArray(record.servedOn) ||
        record.servedOn.some((entry) => typeof entry !== "string")
      )
        throw new AgentDocxError(
          "INVALID_ARGUMENT",
          "service certificate servedOn must be an array of strings",
        );
      if (record.date !== undefined)
        assertDate(
          requiredText(record, "date", "service certificate"),
          "service certificate.date",
        );
    } else if (record.kind === "compliance") {
      assertKeys(
        record,
        ["id", "kind", "basis", "signerCounselId"],
        "compliance certificate",
      );
      requiredDocumentId(record, "id");
      requiredDocumentId(record, "signerCounselId");
      requiredEnum(record, "basis", ["words", "monospaced-lines"]);
    } else
      throw new AgentDocxError(
        "INVALID_ARGUMENT",
        "metadata certificate kind must be service or compliance",
      );
  }
};

export const assertChrome = (value: unknown): void => {
  const chrome = assertRecord(value, "chrome");
  assertKeys(
    chrome,
    ["headers", "footers", "pageNumber", "lineNumbers"],
    "chrome",
  );
  for (const key of ["headers", "footers"] as const) {
    if (chrome[key] === undefined) continue;
    const stories = assertRecord(chrome[key], `chrome.${key}`);
    assertKeys(stories, ["default", "first", "even"], `chrome.${key}`);
    for (const story of ["default", "first", "even"] as const)
      optionalText(stories, story, `chrome.${key}`);
  }
  if (chrome.pageNumber !== undefined) {
    const pageNumber = assertRecord(chrome.pageNumber, "chrome.pageNumber");
    assertKeys(
      pageNumber,
      ["story", "alignment", "format", "start"],
      "chrome.pageNumber",
    );
    requiredEnum(pageNumber, "story", ["header", "footer"]);
    requiredEnum(pageNumber, "alignment", ["left", "center", "right"]);
    requiredEnum(pageNumber, "format", [
      "decimal",
      "lower-roman",
      "upper-roman",
    ]);
    requiredInteger(pageNumber, "start", 1);
  }
  if (chrome.lineNumbers !== undefined) {
    const lineNumbers = assertRecord(chrome.lineNumbers, "chrome.lineNumbers");
    assertKeys(
      lineNumbers,
      ["countBy", "start", "distanceTwips", "restart"],
      "chrome.lineNumbers",
    );
    requiredInteger(lineNumbers, "countBy", 1);
    requiredInteger(lineNumbers, "start", 1);
    requiredInteger(lineNumbers, "distanceTwips", 0);
    requiredEnum(lineNumbers, "restart", [
      "continuous",
      "new-page",
      "new-section",
    ]);
  }
};

export const assertFontSet = (value: unknown): void => {
  const fontSet = assertRecord(value, "fontSet");
  assertKeys(
    fontSet,
    ["family", "regularPath", "boldPath", "italicPath", "boldItalicPath"],
    "fontSet",
  );
  requiredString(fontSet, "family");
  requiredString(fontSet, "regularPath");
  for (const key of ["boldPath", "italicPath", "boldItalicPath"] as const)
    optionalString(fontSet, key);
};

export const assertRendererOptions = (
  value: unknown,
  label: string,
  allowed: readonly string[],
): void => {
  const options = assertRecord(value, label);
  assertKeys(options, allowed, label);
  if (options.renderer !== undefined)
    requiredEnum(options, "renderer", rendererModes);
  if (options.officeTimeoutMs !== undefined)
    optionalInteger(options, "officeTimeoutMs", 1000, 600000);
  if (options.paragraphDiagnostics !== undefined)
    optionalBoolean(options, "paragraphDiagnostics");
  if (options.sectionDiagnostics !== undefined)
    optionalBoolean(options, "sectionDiagnostics");
  if (options.lineDiagnostics !== undefined)
    optionalBoolean(options, "lineDiagnostics");
  if (options.trim !== undefined) {
    if (options.trim !== false) {
      const trim = assertRecord(options.trim, `${label}.trim`);
      assertKeys(trim, ["maxCandidates", "maxLastLineRatio"], `${label}.trim`);
      optionalInteger(trim, "maxCandidates", 1);
      if (
        trim.maxLastLineRatio !== undefined &&
        (typeof trim.maxLastLineRatio !== "number" ||
          !Number.isFinite(trim.maxLastLineRatio) ||
          trim.maxLastLineRatio < 0 ||
          trim.maxLastLineRatio > 1)
      )
        throw new AgentDocxError(
          "INVALID_ARGUMENT",
          `${label}.trim.maxLastLineRatio must be a number from 0 through 1`,
        );
    }
  }
  if (options.word !== undefined) {
    const word = assertRecord(options.word, `${label}.word`);
    assertKeys(word, ["powerShellPath"], `${label}.word`);
    optionalString(word, "powerShellPath");
  }
  if (options.libreoffice !== undefined) {
    const libreoffice = assertRecord(
      options.libreoffice,
      `${label}.libreoffice`,
    );
    assertKeys(
      libreoffice,
      ["executablePath", "installedFonts"],
      `${label}.libreoffice`,
    );
    optionalString(libreoffice, "executablePath");
    if (libreoffice.installedFonts !== undefined) {
      if (!Array.isArray(libreoffice.installedFonts))
        throw new AgentDocxError(
          "INVALID_ARGUMENT",
          `${label}.libreoffice.installedFonts must be an array`,
        );
      for (const font of libreoffice.installedFonts) {
        const installed = assertRecord(
          font,
          `${label}.libreoffice.installedFonts[]`,
        );
        assertKeys(
          installed,
          ["family", "path"],
          `${label}.libreoffice.installedFonts[]`,
        );
        requiredString(installed, "family");
        requiredString(installed, "path");
      }
    }
  }
};

export const assertConfigUpdate = (value: unknown): void => {
  const changes = assertRecord(value, "changes");
  assertKeys(
    changes,
    [
      "profile",
      "filingKind",
      "rulePack",
      "rulePacks",
      "template",
      "assetsDir",
      "fontSet",
      "metadata",
      "chrome",
    ],
    "changes",
  );
  if (Object.keys(changes).length === 0)
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      "changes must include at least one property",
    );
  if (changes.profile !== undefined)
    requiredEnum(changes, "profile", profileIds);
  if (changes.filingKind !== undefined && changes.filingKind !== null)
    requiredEnum(changes, "filingKind", filingKinds);
  if (changes.rulePack !== undefined && changes.rulePack !== null)
    requiredEnum(changes, "rulePack", rulePackIds);
  if (changes.rulePacks !== undefined && changes.rulePacks !== null) {
    if (
      !Array.isArray(changes.rulePacks) ||
      changes.rulePacks.length === 0 ||
      changes.rulePacks.some(
        (entry) => typeof entry !== "string" || entry.length === 0,
      )
    )
      throw new AgentDocxError(
        "INVALID_ARGUMENT",
        "changes.rulePacks must be a non-empty array of pack file paths",
      );
  }
  for (const key of ["template", "assetsDir"] as const)
    if (changes[key] !== undefined && changes[key] !== null)
      requiredString(changes, key);
  if (changes.fontSet !== undefined && changes.fontSet !== null)
    assertFontSet(changes.fontSet);
  if (changes.metadata !== undefined) assertMetadata(changes.metadata);
  if (changes.chrome !== undefined && changes.chrome !== null)
    assertChrome(changes.chrome);
};

export const assertProjectInput = (
  value: Record<string, unknown>,
  includeDefault: boolean,
): void => {
  assertKeys(
    value,
    [
      "documentId",
      "source",
      "createSource",
      "profile",
      "filingKind",
      "rulePack",
      "rulePacks",
      "template",
      "assetsDir",
      "fontSet",
      "metadata",
      "chrome",
      ...(includeDefault ? ["makeDefault"] : []),
    ],
    "project document parameters",
  );
  requiredDocumentId(value, "documentId");
  requiredString(value, "source");
  requiredEnum(value, "profile", profileIds);
  assertMetadata(value.metadata);
  optionalBoolean(value, "createSource");
  optionalEnum(value, "filingKind", filingKinds);
  optionalEnum(value, "rulePack", rulePackIds);
  optionalRulePacks(value);
  optionalString(value, "template");
  optionalString(value, "assetsDir");
  if (value.fontSet !== undefined) assertFontSet(value.fontSet);
  if (value.chrome !== undefined) assertChrome(value.chrome);
  if (includeDefault) optionalBoolean(value, "makeDefault");
};

export const assertSourcePatch = (value: unknown): void => {
  const patch = assertRecord(value, "patch");
  assertKeys(
    patch,
    ["schemaVersion", "documentId", "baseRevision", "edits"],
    "patch",
  );
  if (patch.schemaVersion !== 1)
    throw new AgentDocxError("PATCH_INVALID", "patch.schemaVersion must be 1");
  requiredDocumentId(patch, "documentId");
  requiredRevisionId(patch, "baseRevision");
  if (!Array.isArray(patch.edits))
    throw new AgentDocxError("PATCH_INVALID", "patch.edits must be an array");
  for (const edit of patch.edits) {
    const record = assertRecord(edit, "patch edit");
    assertKeys(
      record,
      ["start", "deleteCount", "expectedText", "replacement"],
      "patch edit",
    );
    requiredInteger(record, "start");
    requiredInteger(record, "deleteCount");
    if ((record.deleteCount as number) < 0)
      throw new AgentDocxError(
        "PATCH_INVALID",
        "patch edit deleteCount must not be negative",
      );
    requiredText(record, "expectedText", "patch edit");
    requiredText(record, "replacement", "patch edit");
  }
};

export const assertChangeSet = (value: unknown): void => {
  const changeSet = assertRecord(value, "changeSet");
  assertKeys(
    changeSet,
    [
      "schemaVersion",
      "id",
      "documentId",
      "baseRevision",
      "headRevision",
      "changes",
      "annotations",
    ],
    "changeSet",
  );
  if (changeSet.schemaVersion !== 1)
    throw new AgentDocxError(
      "CHANGESET_INVALID",
      "changeSet.schemaVersion must be 1",
    );
  requiredRevisionId(changeSet, "id");
  requiredDocumentId(changeSet, "documentId");
  requiredRevisionId(changeSet, "baseRevision");
  requiredRevisionId(changeSet, "headRevision");
  if (
    !Array.isArray(changeSet.changes) ||
    !Array.isArray(changeSet.annotations)
  )
    throw new AgentDocxError(
      "CHANGESET_INVALID",
      "changeSet changes and annotations must be arrays",
    );
  for (const [index, change] of changeSet.changes.entries())
    assertChangeItem(change, `changeSet.changes[${index}]`);
  for (const [index, annotation] of changeSet.annotations.entries())
    assertAnnotationChange(annotation, `changeSet.annotations[${index}]`);
};

const changeKinds = new Set([
  "insert-block",
  "delete-block",
  "move-block",
  "replace-block",
  "replace-container-shell",
  "insert-text",
  "delete-text",
  "replace-text",
  "add-config",
  "remove-config",
  "replace-config",
  "add-dependency",
  "remove-dependency",
  "replace-dependency",
]);

const dateTimePattern =
  /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})(?:\.\d+)?(Z|[+-](\d{2}):(\d{2}))$/;

const assertDateTime = (value: string, label: string): void => {
  const match = dateTimePattern.exec(value);
  const year = match ? Number(match[1]) : NaN;
  const month = match ? Number(match[2]) : NaN;
  const day = match ? Number(match[3]) : NaN;
  const hour = match ? Number(match[4]) : NaN;
  const minute = match ? Number(match[5]) : NaN;
  const second = match ? Number(match[6]) : NaN;
  const offsetHour = match?.[8] === undefined ? 0 : Number(match[8]);
  const offsetMinute = match?.[9] === undefined ? 0 : Number(match[9]);
  const leapYear = year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0);
  const daysInMonth =
    month === 2
      ? leapYear
        ? 29
        : 28
      : [4, 6, 9, 11].includes(month)
        ? 30
        : 31;
  if (
    !match ||
    month < 1 ||
    month > 12 ||
    day < 1 ||
    day > daysInMonth ||
    hour > 23 ||
    minute > 59 ||
    second > 59 ||
    offsetHour > 23 ||
    offsetMinute > 59
  )
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `${label} must be an RFC 3339 date-time or null`,
    );
};
const assertChangeSetActor = (value: unknown, label: string): void => {
  const author = assertRecord(value, label);
  assertKeys(author, ["name", "email"], label);
  if (typeof author.name !== "string" || author.name.length === 0)
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `${label}.name must be a non-empty string`,
    );
  if (author.email !== undefined && typeof author.email !== "string")
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `${label}.email must be a string`,
    );
};

export const assertAttribution = (value: unknown, label: string): void => {
  const attribution = assertRecord(value, `${label}.attribution`);
  assertKeys(
    attribution,
    ["author", "createdAt", "sourceRevisionId"],
    `${label}.attribution`,
  );
  if (!("author" in attribution) || !("createdAt" in attribution))
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `${label}.attribution must have author and createdAt`,
    );
  if (attribution.author !== null)
    assertChangeSetActor(attribution.author, `${label}.attribution.author`);
  if (attribution.createdAt !== null) {
    if (typeof attribution.createdAt !== "string")
      throw new AgentDocxError(
        "INVALID_ARGUMENT",
        `${label}.attribution.createdAt must be a date-time or null`,
      );
    assertDateTime(attribution.createdAt, `${label}.attribution.createdAt`);
  }
  if (
    attribution.sourceRevisionId !== undefined &&
    typeof attribution.sourceRevisionId !== "string"
  )
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `${label}.attribution.sourceRevisionId must be a string`,
    );
};

export const assertBlockLocation = (value: unknown, label: string): void => {
  const location = assertRecord(value, label);
  assertKeys(
    location,
    ["collection", "parentId", "index", "sourceOffset"],
    label,
  );
  if (location.collection !== "body" && location.collection !== "footnotes")
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `${label}.collection must be body or footnotes`,
    );
  if (
    location.parentId !== null &&
    (typeof location.parentId !== "string" ||
      !blockIdPattern.test(location.parentId))
  )
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `${label}.parentId must be a block ID or null`,
    );
  requiredInteger(location, "index");
  requiredInteger(location, "sourceOffset");
};

export const assertSourceRange = (value: unknown, label: string): void => {
  const range = assertRecord(value, label);
  assertKeys(range, ["start", "end", "text"], label);
  requiredInteger(range, "start");
  requiredInteger(range, "end");
  if ((range.end as number) < (range.start as number))
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `${label}.end must not precede start`,
    );
  requiredText(range, "text", label);
};
const assertAttributionSpans = (value: unknown, label: string): void => {
  if (!Array.isArray(value))
    throw new AgentDocxError("INVALID_ARGUMENT", `${label} must be an array`);
  for (const [index, entry] of value.entries()) {
    const span = assertRecord(entry, `${label}[${index}]`);
    assertKeys(span, ["start", "end", "attribution"], `${label}[${index}]`);
    const start = requiredInteger(span, "start");
    const end = requiredInteger(span, "end");
    if (end < start)
      throw new AgentDocxError(
        "INVALID_ARGUMENT",
        `${label}[${index}].end must not precede start`,
      );
    assertAttribution(span.attribution, `${label}[${index}]`);
  }
};

export const assertChangeItem = (value: unknown, label: string): void => {
  const change = assertRecord(value, label);
  if (
    Object.keys(change).some(
      (key) =>
        ![
          "id",
          "kind",
          "attribution",
          "blockId",
          "from",
          "to",
          "oldSource",
          "newSource",
          "oldOffset",
          "newOffset",
          "oldText",
          "newText",
          "oldAttributionSpans",
          "newAttributionSpans",
          "block",
          "oldBlock",
          "newBlock",
          "oldShell",
          "newShell",
          "path",
          "oldValue",
          "newValue",
          "key",
          "oldObject",
          "newObject",
        ].includes(key),
    )
  )
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `${label} has an unsupported property`,
    );
  if (typeof change.id !== "string" || !changeIdPattern.test(change.id))
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `${label}.id must be a canonical change ID`,
    );
  if (typeof change.kind !== "string" || !changeKinds.has(change.kind))
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `${label}.kind is not a supported change kind`,
    );
  assertAttribution(change.attribution, label);
  if (change.blockId !== undefined) {
    if (
      typeof change.blockId !== "string" ||
      !blockIdPattern.test(change.blockId)
    )
      throw new AgentDocxError(
        "INVALID_ARGUMENT",
        `${label}.blockId must be a block ID`,
      );
  }
  if (change.from !== undefined)
    assertBlockLocation(change.from, `${label}.from`);
  if (change.to !== undefined) assertBlockLocation(change.to, `${label}.to`);
  if (change.oldSource !== undefined)
    assertSourceRange(change.oldSource, `${label}.oldSource`);
  if (change.newSource !== undefined)
    assertSourceRange(change.newSource, `${label}.newSource`);
  for (const key of ["oldAttributionSpans", "newAttributionSpans"] as const)
    if (change[key] !== undefined)
      assertAttributionSpans(change[key], `${label}.${key}`);
  for (const key of ["oldOffset", "newOffset"] as const)
    if (change[key] !== undefined) requiredInteger(change, key);
};

export const assertAnnotationChange = (value: unknown, label: string): void => {
  const change = assertRecord(value, label);
  assertKeys(change, ["id", "kind", "oldValue", "newValue"], label);
  if (typeof change.id !== "string" || !changeIdPattern.test(change.id))
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `${label}.id must be a canonical change ID`,
    );
  if (
    change.kind !== "add" &&
    change.kind !== "replace" &&
    change.kind !== "remove"
  )
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `${label}.kind must be add, replace, or remove`,
    );
  if (change.kind === "add" && change.newValue === undefined)
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `${label} must carry newValue`,
    );
  if (change.kind === "remove" && change.oldValue === undefined)
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `${label} must carry oldValue`,
    );
  for (const key of ["oldValue", "newValue"] as const)
    if (change[key] !== undefined)
      assertReviewAnnotation(change[key]!, `${label}.${key}`);
};

export const assertReviewAnnotation = (value: unknown, label: string): void => {
  const annotation = assertRecord(value, label);
  assertKeys(
    annotation,
    ["id", "blockId", "range", "author", "createdAt", "message", "status"],
    label,
  );
  for (const key of [
    "id",
    "blockId",
    "author",
    "createdAt",
    "message",
    "status",
  ])
    if (!(key in annotation))
      throw new AgentDocxError("INVALID_ARGUMENT", `${label} must have ${key}`);
  if (
    typeof annotation.id !== "string" ||
    !annotationIdPattern.test(annotation.id)
  )
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `${label}.id must be an annotation ID`,
    );
  if (
    typeof annotation.blockId !== "string" ||
    !blockIdPattern.test(annotation.blockId)
  )
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `${label}.blockId must be a block ID`,
    );
  if (annotation.range !== undefined) {
    const range = assertRecord(annotation.range, `${label}.range`);
    assertKeys(range, ["start", "end"], `${label}.range`);
    const start = requiredInteger(range, "start");
    const end = requiredInteger(range, "end");
    if (end < start)
      throw new AgentDocxError(
        "INVALID_ARGUMENT",
        `${label}.range.end must not precede start`,
      );
  }
  if (annotation.author !== null)
    assertChangeSetActor(annotation.author, `${label}.author`);
  if (annotation.createdAt !== null) {
    if (typeof annotation.createdAt !== "string")
      throw new AgentDocxError(
        "INVALID_ARGUMENT",
        `${label}.createdAt must be a date-time or null`,
      );
    assertDateTime(annotation.createdAt, `${label}.createdAt`);
  }
  requiredText(annotation, "message", label);
  if (annotation.status !== "open" && annotation.status !== "resolved")
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `${label}.status must be open or resolved`,
    );
};

export const assertResolutionDecisions = (value: unknown): void => {
  const decisions = assertRecord(value, "decisions");
  for (const [id, decision] of Object.entries(decisions)) {
    if (
      !changeIdPattern.test(id) ||
      (decision !== "accept" && decision !== "reject")
    )
      throw new AgentDocxError(
        "CHANGESET_INVALID",
        "Resolution decisions must map canonical change IDs to accept or reject",
      );
  }
};

export const assertReviewRange = (value: unknown): void => {
  const range = assertRecord(value, "range");
  assertKeys(range, ["start", "length"], "range");
  const start = requiredInteger(range, "start");
  const length = requiredInteger(range, "length");
  if (length < 0)
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      "Review range length must not be negative",
    );
  void start;
};

export const assertFilingSetId = (value: Record<string, unknown>): void => {
  if (typeof value.id !== "string" || !documentIdPattern.test(value.id))
    throw new AgentDocxError("INVALID_ARGUMENT", "id must be a filing set ID");
};

export const assertImportParams = (value: Record<string, unknown>): void => {
  assertKeys(
    value,
    [
      "input",
      "attachments",
      "inspectOnly",
      "documentId",
      "output",
      "author",
      "message",
    ],
    "docx.import params",
  );
  requiredString(value, "input");
  optionalString(value, "attachments");
  if (typeof value.inspectOnly !== "boolean")
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      "inspectOnly must be a boolean",
    );
  if (value.inspectOnly) {
    for (const key of ["documentId", "output", "author", "message"] as const)
      if (hasOwn(value, key))
        throw new AgentDocxError(
          "INVALID_ARGUMENT",
          "Inspect-only DOCX import is stateless",
        );
    return;
  }
  requiredDocumentId(value, "documentId");
  requiredString(value, "output");
  actor(value.author);
  requiredString(value, "message");
};

export const isStatelessAgentRequest = (
  action: AgentAction,
  params: Record<string, unknown>,
): boolean =>
  action === "docx.inspect" ||
  (action === "docx.import" && params.inspectOnly === true);

const assertAgentParams = (
  action: AgentAction,
  params: Record<string, unknown>,
): void => {
  switch (action) {
    case "project.init":
      assertProjectInput(params, false);
      return;
    case "project.add":
      assertProjectInput(params, true);
      return;
    case "project.get":
      noOptions(params, "project.get params");
      return;
    case "document.configure":
      assertKeys(
        params,
        ["documentId", "baseRevision", "changes", "author", "message"],
        "document.configure params",
      );
      requiredDocumentId(params, "documentId");
      asRevision(params.baseRevision, "baseRevision", true);
      assertConfigUpdate(params.changes);
      actor(params.author);
      requiredString(params, "message");
      return;
    case "document.get":
    case "document.validate":
    case "draft.guidance":
      assertKeys(params, ["documentId", "revision"], `${action} params`);
      requiredDocumentId(params, "documentId");
      optionalRevision(params, "revision");
      return;
    case "document.measure":
      assertKeys(
        params,
        ["documentId", "revision", "options"],
        "document.measure params",
      );
      if (params.options !== undefined)
        assertRendererOptions(params.options, "document.measure options", [
          "renderer",
          "officeTimeoutMs",
          "paragraphDiagnostics",
          "sectionDiagnostics",
          "lineDiagnostics",
          "trim",
          "word",
          "libreoffice",
        ]);
      return;
    case "revision.checkpoint":
      assertKeys(
        params,
        ["documentId", "baseRevision", "author", "message"],
        "revision.checkpoint params",
      );
      requiredDocumentId(params, "documentId");
      asRevision(params.baseRevision, "baseRevision", true);
      actor(params.author);
      requiredString(params, "message");
      return;
    case "revision.list":
      assertKeys(
        params,
        ["documentId", "limit", "cursor"],
        "revision.list params",
      );
      requiredDocumentId(params, "documentId");
      optionalInteger(params, "limit", 1, 1000);
      if (params.cursor !== undefined) requiredRevisionId(params, "cursor");
      return;
    case "revision.get":
      assertKeys(params, ["documentId", "revision"], "revision.get params");
      requiredDocumentId(params, "documentId");
      asRevision(params.revision, "revision");
      return;
    case "revision.restore":
      assertKeys(
        params,
        ["documentId", "baseRevision", "targetRevision", "author", "message"],
        "revision.restore params",
      );
      requiredDocumentId(params, "documentId");
      asRevision(params.baseRevision, "baseRevision");
      asRevision(params.targetRevision, "targetRevision");
      actor(params.author);
      requiredString(params, "message");
      return;
    case "revision.diff":
      assertKeys(
        params,
        ["documentId", "baseRevision", "headRevision", "output"],
        "revision.diff params",
      );
      requiredDocumentId(params, "documentId");
      asRevision(params.baseRevision, "baseRevision");
      asRevision(params.headRevision, "headRevision");
      optionalString(params, "output");
      return;
    case "revision.resolve":
      assertKeys(
        params,
        ["documentId", "changeSet", "decisions", "author", "message"],
        "revision.resolve params",
      );
      requiredDocumentId(params, "documentId");
      assertChangeSet(params.changeSet);
      assertResolutionDecisions(params.decisions);
      actor(params.author);
      requiredString(params, "message");
      return;
    case "draft.evaluate":
      assertKeys(params, ["patch", "renderer"], "draft.evaluate params");
      assertSourcePatch(params.patch);
      optionalEnum(params, "renderer", rendererModes);
      return;
    case "draft.apply":
      assertKeys(
        params,
        ["patch", "patchHash", "gate", "author", "message"],
        "draft.apply params",
      );
      assertSourcePatch(params.patch);
      requiredRevisionId(params, "patchHash");
      optionalEnum(params, "gate", ["report", "not-worse", "pass"]);
      actor(params.author);
      requiredString(params, "message");
      return;
    case "review.add":
      assertKeys(
        params,
        ["documentId", "revision", "blockId", "range", "author", "message"],
        "review.add params",
      );
      requiredDocumentId(params, "documentId");
      asRevision(params.revision, "revision");
      if (
        typeof params.blockId !== "string" ||
        !blockIdPattern.test(params.blockId)
      )
        throw new AgentDocxError(
          "INVALID_ARGUMENT",
          "blockId must be a block ID",
        );
      if (params.range !== undefined) assertReviewRange(params.range);
      actor(params.author);
      requiredString(params, "message");
      return;
    case "review.resolve":
      assertKeys(
        params,
        ["documentId", "revision", "annotationId", "author", "message"],
        "review.resolve params",
      );
      requiredDocumentId(params, "documentId");
      asRevision(params.revision, "revision");
      if (
        typeof params.annotationId !== "string" ||
        !annotationIdPattern.test(params.annotationId)
      )
        throw new AgentDocxError(
          "INVALID_ARGUMENT",
          "annotationId must be an annotation ID",
        );
      actor(params.author);
      requiredString(params, "message");
      return;
    case "docx.export": {
      assertKeys(
        params,
        ["documentId", "revision", "mode", "baseRevision", "output", "options"],
        "docx.export params",
      );
      requiredDocumentId(params, "documentId");
      asRevision(params.revision, "revision");
      const mode = requiredEnum(params, "mode", ["clean", "redline", "pdf"]);
      requiredString(params, "output");
      if (
        (mode === "clean" || mode === "pdf") &&
        hasOwn(params, "baseRevision")
      )
        throw new AgentDocxError(
          "INVALID_ARGUMENT",
          `${mode === "clean" ? "Clean" : "PDF"} DOCX export forbids baseRevision`,
        );
      if (mode === "redline") asRevision(params.baseRevision, "baseRevision");
      if (params.options !== undefined)
        assertRendererOptions(params.options, "docx.export options", [
          "renderer",
          "officeTimeoutMs",
          "word",
          "libreoffice",
        ]);
      return;
    }
    case "docx.importRedline":
      assertKeys(
        params,
        ["documentId", "input", "attachments", "author", "message"],
        "docx.importRedline params",
      );
      requiredDocumentId(params, "documentId");
      requiredString(params, "input");
      optionalString(params, "attachments");
      actor(params.author);
      requiredString(params, "message");
      return;
    case "filingSet.add":
      assertKeys(
        params,
        ["id", "label", "documentIds", "pageCap"],
        "filingSet.add params",
      );
      assertFilingSetId(params);
      optionalString(params, "label");
      if (
        !Array.isArray(params.documentIds) ||
        params.documentIds.length === 0 ||
        params.documentIds.some(
          (entry) =>
            typeof entry !== "string" || !documentIdPattern.test(entry),
        )
      )
        throw new AgentDocxError(
          "INVALID_ARGUMENT",
          "documentIds must be a non-empty array of document IDs",
        );
      const seen: Record<string, true> = {};
      for (const entry of params.documentIds as string[]) {
        if (seen[entry])
          throw new AgentDocxError(
            "INVALID_ARGUMENT",
            "documentIds must not contain duplicates",
          );
        seen[entry] = true;
      }
      optionalInteger(params, "pageCap", 1);
      return;
    case "filingSet.remove":
    case "filingSet.get":
    case "filingSet.validate":
      assertKeys(params, ["id"], `${action} params`);
      assertFilingSetId(params);
      return;
    case "docx.import":
      assertImportParams(params);
      return;
    case "docx.inspect":
      assertKeys(params, ["input"], "docx.inspect params");
      requiredString(params, "input");
      return;
  }
};

export const parseAgentRequest = (value: unknown): AgentRequest => {
  const request = assertRecord(value, "Agent request");
  assertKeys(
    request,
    ["schemaVersion", "id", "action", "project", "params"],
    "Agent request",
  );
  if (request.schemaVersion !== 1)
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      "Agent request schemaVersion must be 1",
    );
  if (
    hasOwn(request, "id") &&
    request.id !== null &&
    typeof request.id !== "string" &&
    (typeof request.id !== "number" || !Number.isFinite(request.id))
  )
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      "Agent request id must be a string, finite number, or null",
    );
  if (
    typeof request.action !== "string" ||
    !(agentActions as readonly string[]).includes(request.action)
  )
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      "Agent request action is invalid",
    );
  if (
    hasOwn(request, "project") &&
    (typeof request.project !== "string" || request.project === "")
  )
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      "Agent request project must be a non-empty path",
    );
  const action = request.action as AgentAction;
  const params = assertRecord(request.params, "Agent request params");
  assertAgentParams(action, params);
  if (hasOwn(request, "project") && isStatelessAgentRequest(action, params))
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `${action} is stateless and forbids project`,
    );
  return {
    schemaVersion: 1,
    ...(hasOwn(request, "id") ? { id: request.id as AgentRequestId } : {}),
    action,
    ...(hasOwn(request, "project")
      ? { project: request.project as string }
      : {}),
    params: params as AgentParamsForAction<typeof action>,
  } as AgentRequest;
};
