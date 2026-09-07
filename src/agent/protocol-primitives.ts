import { dirname, isAbsolute, relative, resolve, sep } from "node:path";
import { publicPath } from "../path-util.js";
import type {
  DocumentConfigUpdate,
  ResolveChangesInput,
} from "../project/contracts.js";
import type { Actor, RevisionId } from "../legal/model.js";
import type { ChangeSet } from "../revisions/types.js";
import type { SourcePatch } from "../draft/types.js";
import { AgentDocxError } from "../types.js";
import { objectRecord } from "../json-contract.js";

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

export const hasOwn = (value: Record<string, unknown>, key: string): boolean =>
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

export const assertDate = (value: string, label: string): void => {
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
