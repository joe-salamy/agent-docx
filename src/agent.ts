import { readFile } from "node:fs/promises";
import { dirname, isAbsolute, relative, resolve, sep } from "node:path";
import { inspectDocxTemplate } from "./docx/inspect.js";
import { inspectDocx } from "./docx/import.js";
import type { ExportDocxInput, ImportDocxInput } from "./docx/contracts.js";
import type {
  AddReviewInput,
  ConfigureDocumentInput,
  DocumentConfigUpdate,
  ProjectDocumentInput,
  ResolveChangesInput,
  ResolveReviewInput,
} from "./project/contracts.js";
import { createProject, openProject } from "./project/index.js";
import type { Actor, RevisionId } from "./legal/model.js";
import type { ChangeSet } from "./revisions/types.js";
import type { SourcePatch } from "./draft/types.js";
import { AgentDocxError, type JsonValue } from "./types.js";
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
export type AgentRequest = {
  schemaVersion: 1;
  id?: string | number | null;
  action: AgentAction;
  project?: string;
  params: Record<string, unknown>;
};

export type AgentDispatchResult = {
  request: AgentRequest;
  project: string | null;
  documentId: string | null;
  revision: RevisionId | null;
  value: unknown;
};

const assertRecord = (value: unknown, label: string): Record<string, unknown> =>
  objectRecord(value, label);

const assertKeys = (
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

const requiredString = (
  value: Record<string, unknown>,
  key: string,
): string => {
  if (typeof value[key] !== "string" || value[key] === "")
    throw new AgentDocxError("INVALID_ARGUMENT", `${key} is required`);
  return value[key] as string;
};

const optionalString = (
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

const optionalBoolean = (
  value: Record<string, unknown>,
  key: string,
): boolean | undefined => {
  if (value[key] === undefined) return undefined;
  if (typeof value[key] !== "boolean")
    throw new AgentDocxError("INVALID_ARGUMENT", `${key} must be a boolean`);
  return value[key] as boolean;
};

const asRevision = (
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

const actor = (value: unknown): Actor => {
  const candidate = assertRecord(value, "author");
  assertKeys(candidate, ["name", "email"], "author");
  const name = requiredString(candidate, "name");
  const email = optionalString(candidate, "email");
  return email === undefined ? { name } : { name, email };
};

const invocationPath = (cwd: string, path: string): string =>
  isAbsolute(path) ? path : resolve(cwd, path);

const publicPath = (cwd: string, path: string): string => {
  const output = relative(cwd, isAbsolute(path) ? path : resolve(cwd, path))
    .split(sep)
    .join("/");
  return output === "" ? "." : output;
};

const manifestRelativePath = (
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
    throw new AgentDocxError(
      "PATH_OUTSIDE_PROJECT",
      `Path is outside project: ${path}`,
    );
  return candidate;
};

const projectPath = (cwd: string, requested: string | undefined): string =>
  invocationPath(cwd, requested ?? "agent-docx.json");

const projectInput = (
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
    profile: profile as ProjectDocumentInput["profile"],
    metadata: metadata as ProjectDocumentInput["metadata"],
    ...(createSource === undefined ? {} : { createSource }),
    ...(template
      ? { template: manifestRelativePath(cwd, manifestPath, template) }
      : {}),
    ...(assetsDir
      ? { assetsDir: manifestRelativePath(cwd, manifestPath, assetsDir) }
      : {}),
    ...(filingKind
      ? { filingKind: filingKind as ProjectDocumentInput["filingKind"] }
      : {}),
    ...(rulePack
      ? { rulePack: rulePack as ProjectDocumentInput["rulePack"] }
      : {}),
    ...(rulePacks
      ? {
          rulePacks: rulePacks.map((path) =>
            manifestRelativePath(cwd, manifestPath, path),
          ),
        }
      : {}),
    ...(fontSet ? { fontSet } : {}),
    ...(chrome ? { chrome: chrome as ProjectDocumentInput["chrome"] } : {}),
    ...(includeDefault && optionalBoolean(params, "makeDefault")
      ? { makeDefault: true }
      : {}),
  };
};
const configUpdate = (
  cwd: string,
  manifestPath: string,
  value: unknown,
): DocumentConfigUpdate => {
  assertConfigUpdate(value);
  const raw = assertRecord(value, "changes");
  const changes: DocumentConfigUpdate = {};
  if (raw.profile !== undefined)
    changes.profile = raw.profile as DocumentConfigUpdate["profile"];
  if (raw.filingKind !== undefined)
    changes.filingKind = raw.filingKind as DocumentConfigUpdate["filingKind"];
  if (raw.rulePack !== undefined)
    changes.rulePack = raw.rulePack as DocumentConfigUpdate["rulePack"];
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
    changes.metadata = raw.metadata as DocumentConfigUpdate["metadata"];
  if (raw.chrome !== undefined)
    changes.chrome = raw.chrome as DocumentConfigUpdate["chrome"];
  return changes;
};

const asPatch = (value: unknown): SourcePatch => {
  assertSourcePatch(value);
  return value as SourcePatch;
};

const asChangeSet = (value: unknown): ChangeSet => {
  assertChangeSet(value);
  return value as ChangeSet;
};

const noOptions = (value: unknown, label: string): Record<string, never> => {
  const result = assertRecord(value, label);
  assertKeys(result, [], label);
  return result as Record<string, never>;
};
const documentIdPattern = /^[a-z0-9][a-z0-9-]{0,63}$/;
const blockIdPattern =
  /^b_[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/;
const annotationIdPattern =
  /^a_[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/;
const revisionIdPattern = /^sha256:[0-9a-f]{64}$/;
const changeIdPattern = /^c_[0-9a-f]{64}$/;

const profileIds = [
  "us-district-conventional",
  "frap-32",
  "cand-civil",
] as const;
const filingKinds = [
  "principal-brief",
  "reply-brief",
  "motion-document",
  "opposition-text",
  "reply-text",
] as const;
const rulePackIds = ["frap-32@2024-12-01", "cand-civil@2026-05-01"] as const;
const rendererModes = [
  "deterministic",
  "word",
  "libreoffice",
  "compare",
] as const;

const hasOwn = (value: Record<string, unknown>, key: string): boolean =>
  Object.prototype.hasOwnProperty.call(value, key);

const requiredDocumentId = (
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

const requiredEnum = <T extends readonly string[]>(
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

const optionalEnum = <T extends readonly string[]>(
  value: Record<string, unknown>,
  key: string,
  allowed: T,
): T[number] | undefined => {
  if (value[key] === undefined) return undefined;
  return requiredEnum(value, key, allowed);
};

const requiredRevisionId = (
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

const optionalRevision = (
  value: Record<string, unknown>,
  key: string,
): RevisionId | "HEAD" | undefined =>
  value[key] === undefined
    ? undefined
    : (asRevision(value[key], key) as RevisionId | "HEAD");

const optionalRulePacks = (
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

const requiredInteger = (
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

const optionalInteger = (
  value: Record<string, unknown>,
  key: string,
  minimum = 0,
  maximum = Number.MAX_SAFE_INTEGER,
): number | undefined =>
  value[key] === undefined
    ? undefined
    : requiredInteger(value, key, minimum, maximum);

const requiredText = (
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

const optionalText = (
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

const assertMetadata = (value: unknown): void => {
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

const assertChrome = (value: unknown): void => {
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

const assertFontSet = (value: unknown): void => {
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

const assertRendererOptions = (
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

const assertConfigUpdate = (value: unknown): void => {
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

const assertProjectInput = (
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

const assertSourcePatch = (value: unknown): void => {
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
      ["start", "end", "expectedText", "replacement"],
      "patch edit",
    );
    requiredInteger(record, "start");
    requiredInteger(record, "end");
    if ((record.end as number) < (record.start as number))
      throw new AgentDocxError(
        "PATCH_INVALID",
        "patch edit end must not precede start",
      );
    requiredText(record, "expectedText", "patch edit");
    requiredText(record, "replacement", "patch edit");
  }
};

const assertChangeSet = (value: unknown): void => {
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
};

const assertResolutionDecisions = (value: unknown): void => {
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

const assertReviewRange = (value: unknown): void => {
  const range = assertRecord(value, "range");
  assertKeys(range, ["start", "end"], "range");
  const start = requiredInteger(range, "start");
  const end = requiredInteger(range, "end");
  if (end < start)
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      "Review range end must not precede start",
    );
};

const assertFilingSetId = (value: Record<string, unknown>): void => {
  if (typeof value.id !== "string" || !documentIdPattern.test(value.id))
    throw new AgentDocxError("INVALID_ARGUMENT", "id must be a filing set ID");
};

const assertImportParams = (value: Record<string, unknown>): void => {
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
      requiredDocumentId(params, "documentId");
      optionalRevision(params, "revision");
      if (params.options !== undefined)
        assertRendererOptions(params.options, "document.measure options", [
          "renderer",
          "officeTimeoutMs",
          "paragraphDiagnostics",
          "sectionDiagnostics",
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
    ...(hasOwn(request, "id")
      ? { id: request.id as string | number | null }
      : {}),
    action,
    ...(hasOwn(request, "project")
      ? { project: request.project as string }
      : {}),
    params,
  };
};

const responseMeta = (
  action: AgentAction,
  params: Record<string, unknown>,
  value: unknown,
): { documentId: string | null; revision: RevisionId | null } => {
  const patch =
    typeof params.patch === "object" &&
    params.patch !== null &&
    !Array.isArray(params.patch)
      ? (params.patch as Record<string, unknown>)
      : null;
  const documentId =
    typeof params.documentId === "string"
      ? params.documentId
      : patch && typeof patch.documentId === "string"
        ? patch.documentId
        : null;
  const record =
    typeof value === "object" && value !== null && !Array.isArray(value)
      ? (value as Record<string, unknown>)
      : null;
  const revisionId = (candidate: unknown): RevisionId | null =>
    typeof candidate === "string" && revisionIdPattern.test(candidate)
      ? (candidate as RevisionId)
      : null;

  if (action === "docx.inspect" || isStatelessAgentRequest(action, params))
    return { documentId: null, revision: null };

  if (action === "project.init" || action === "project.add") {
    const documents = record?.documents;
    const document = Array.isArray(documents)
      ? documents.find(
          (entry) =>
            typeof entry === "object" &&
            entry !== null &&
            !Array.isArray(entry) &&
            (entry as Record<string, unknown>).documentId === documentId,
        )
      : undefined;
    const head =
      typeof document === "object" &&
      document !== null &&
      !Array.isArray(document)
        ? revisionId((document as Record<string, unknown>).head)
        : null;
    return { documentId, revision: head };
  }

  if (action === "project.get") return { documentId: null, revision: null };

  if (action === "document.get")
    return {
      documentId,
      revision: record ? revisionId(record.revision) : null,
    };

  if (action === "revision.list") return { documentId, revision: null };

  if (action === "revision.get")
    return { documentId, revision: record ? revisionId(record.id) : null };

  if (action === "revision.diff") {
    const changeSet =
      record?.changeSet &&
      typeof record.changeSet === "object" &&
      !Array.isArray(record.changeSet)
        ? (record.changeSet as Record<string, unknown>)
        : null;
    return {
      documentId,
      revision: changeSet ? revisionId(changeSet.headRevision) : null,
    };
  }

  if (action === "draft.evaluate") {
    const patch =
      typeof params.patch === "object" &&
      params.patch !== null &&
      !Array.isArray(params.patch)
        ? (params.patch as Record<string, unknown>)
        : null;
    return {
      documentId:
        patch && typeof patch.documentId === "string" ? patch.documentId : null,
      revision: patch ? revisionId(patch.baseRevision) : null,
    };
  }

  if (action === "docx.import")
    return {
      documentId,
      revision: record ? revisionId(record.headRevision) : null,
    };

  if (action === "docx.importRedline")
    return {
      documentId,
      revision: record ? revisionId(record.headRevision) : null,
    };

  if (record) {
    const head = revisionId(record.head);
    if (head) return { documentId, revision: head };

    const revision =
      typeof record.revision === "object" &&
      record.revision !== null &&
      !Array.isArray(record.revision)
        ? (record.revision as Record<string, unknown>)
        : null;
    const mutationRevision = revision ? revisionId(revision.id) : null;
    if (mutationRevision) return { documentId, revision: mutationRevision };

    const selectedRevision = revisionId(record.revision);
    if (selectedRevision) return { documentId, revision: selectedRevision };

    const measurement =
      typeof record.measurement === "object" &&
      record.measurement !== null &&
      !Array.isArray(record.measurement)
        ? (record.measurement as Record<string, unknown>)
        : null;
    const measuredRevision = measurement
      ? revisionId(measurement.revision)
      : null;
    if (measuredRevision) return { documentId, revision: measuredRevision };

    const artifact =
      typeof record.artifact === "object" &&
      record.artifact !== null &&
      !Array.isArray(record.artifact)
        ? (record.artifact as Record<string, unknown>)
        : null;
    const artifactRevision = artifact ? revisionId(artifact.revision) : null;
    if (artifactRevision) return { documentId, revision: artifactRevision };
  }
  return { documentId, revision: null };
};

/** Executes one validated local agent request without exposing binary DOCX bytes. */
export const dispatchAgentRequest = async (
  raw: unknown,
  cwd: string,
): Promise<AgentDispatchResult> => {
  const request = parseAgentRequest(raw);
  const manifestPath = projectPath(cwd, request.project);
  const projectDisplay = isStatelessAgentRequest(request.action, request.params)
    ? null
    : publicPath(cwd, manifestPath);
  let value: unknown;
  switch (request.action) {
    case "project.init": {
      const input = projectInput(cwd, manifestPath, request.params, false);
      value = await (await createProject(manifestPath, input)).getState();
      break;
    }
    case "project.add": {
      const project = await openProject(manifestPath);
      value = await project.addDocument(
        projectInput(cwd, manifestPath, request.params, true),
      );
      break;
    }
    case "project.get": {
      noOptions(request.params, "project.get params");
      value = await (await openProject(manifestPath)).getState();
      break;
    }
    case "document.configure": {
      const changes = configUpdate(cwd, manifestPath, request.params.changes);
      const input: ConfigureDocumentInput = {
        baseRevision: asRevision(
          request.params.baseRevision,
          "baseRevision",
          true,
        ),
        changes,
        author: actor(request.params.author),
        message: requiredString(request.params, "message"),
      };
      value = await (
        await openProject(manifestPath)
      ).configureDocument(
        requiredDocumentId(request.params, "documentId"),
        input,
      );
      break;
    }
    case "document.get": {
      assertKeys(
        request.params,
        ["documentId", "revision"],
        "document.get params",
      );
      const revision =
        request.params.revision === undefined
          ? undefined
          : (asRevision(request.params.revision, "revision") as
              | RevisionId
              | "HEAD");
      value = await (
        await openProject(manifestPath)
      ).getDocument(requiredString(request.params, "documentId"), revision);
      break;
    }
    case "document.measure": {
      assertKeys(
        request.params,
        ["documentId", "revision", "options"],
        "document.measure params",
      );
      const options =
        request.params.options === undefined
          ? {}
          : assertRecord(request.params.options, "options");
      const revision =
        request.params.revision === undefined
          ? undefined
          : (asRevision(request.params.revision, "revision") as
              | RevisionId
              | "HEAD");
      value = await (
        await openProject(manifestPath)
      ).measure(
        requiredString(request.params, "documentId"),
        revision,
        options,
      );
      break;
    }
    case "document.validate": {
      assertKeys(
        request.params,
        ["documentId", "revision"],
        "document.validate params",
      );
      const revision =
        request.params.revision === undefined
          ? undefined
          : (asRevision(request.params.revision, "revision") as
              | RevisionId
              | "HEAD");
      value = await (
        await openProject(manifestPath)
      ).validate(requiredString(request.params, "documentId"), revision);
      break;
    }
    case "revision.checkpoint": {
      assertKeys(
        request.params,
        ["documentId", "baseRevision", "author", "message"],
        "revision.checkpoint params",
      );
      value = await (
        await openProject(manifestPath)
      ).checkpoint(requiredString(request.params, "documentId"), {
        baseRevision: asRevision(
          request.params.baseRevision,
          "baseRevision",
          true,
        ),
        author: actor(request.params.author),
        message: requiredString(request.params, "message"),
      });
      break;
    }
    case "revision.list": {
      const cursor =
        request.params.cursor === undefined
          ? undefined
          : requiredRevisionId(request.params, "cursor");
      value = await (
        await openProject(manifestPath)
      ).listRevisions(requiredDocumentId(request.params, "documentId"), {
        ...(request.params.limit === undefined
          ? {}
          : { limit: requiredInteger(request.params, "limit", 1, 1000) }),
        ...(cursor === undefined ? {} : { cursor }),
      });
      break;
    }
    case "revision.get": {
      assertKeys(
        request.params,
        ["documentId", "revision"],
        "revision.get params",
      );
      value = await (
        await openProject(manifestPath)
      ).getRevision(
        requiredString(request.params, "documentId"),
        asRevision(request.params.revision, "revision") as RevisionId | "HEAD",
      );
      break;
    }
    case "revision.restore": {
      assertKeys(
        request.params,
        ["documentId", "baseRevision", "targetRevision", "author", "message"],
        "revision.restore params",
      );
      value = await (
        await openProject(manifestPath)
      ).restore(requiredString(request.params, "documentId"), {
        baseRevision: asRevision(
          request.params.baseRevision,
          "baseRevision",
        ) as RevisionId | "HEAD",
        targetRevision: asRevision(
          request.params.targetRevision,
          "targetRevision",
        ) as RevisionId | "HEAD",
        author: actor(request.params.author),
        message: requiredString(request.params, "message"),
      });
      break;
    }
    case "revision.diff": {
      assertKeys(
        request.params,
        ["documentId", "baseRevision", "headRevision", "output"],
        "revision.diff params",
      );
      const project = await openProject(manifestPath);
      const documentId = requiredString(request.params, "documentId");
      const baseRevision = asRevision(
        request.params.baseRevision,
        "baseRevision",
      ) as RevisionId | "HEAD";
      const headRevision = asRevision(
        request.params.headRevision,
        "headRevision",
      ) as RevisionId | "HEAD";
      const changeSet = await project.diff(
        documentId,
        baseRevision,
        headRevision,
      );
      const output = optionalString(request.params, "output");
      value = output
        ? {
            changeSet,
            compiled: await project.exportDocx(documentId, {
              revision: headRevision,
              mode: "redline",
              baseRevision,
              output: invocationPath(cwd, output),
            }),
          }
        : { changeSet };
      break;
    }
    case "revision.resolve": {
      assertKeys(
        request.params,
        ["documentId", "changeSet", "decisions", "author", "message"],
        "revision.resolve params",
      );
      const decisions = assertRecord(request.params.decisions, "decisions");
      for (const decision of Object.values(decisions))
        if (decision !== "accept" && decision !== "reject")
          throw new AgentDocxError(
            "CHANGESET_INVALID",
            "Resolution decisions must be accept or reject",
          );
      const input: ResolveChangesInput = {
        changeSet: asChangeSet(request.params.changeSet),
        decisions: decisions as ResolveChangesInput["decisions"],
        author: actor(request.params.author),
        message: requiredString(request.params, "message"),
      };
      value = await (
        await openProject(manifestPath)
      ).resolveChanges(requiredString(request.params, "documentId"), input);
      break;
    }
    case "draft.guidance": {
      assertKeys(
        request.params,
        ["documentId", "revision"],
        "draft.guidance params",
      );
      const revision =
        request.params.revision === undefined
          ? undefined
          : (asRevision(request.params.revision, "revision") as
              | RevisionId
              | "HEAD");
      value = await (
        await openProject(manifestPath)
      ).getDraftGuidance(
        requiredString(request.params, "documentId"),
        revision,
      );
      break;
    }
    case "draft.evaluate": {
      assertKeys(
        request.params,
        ["patch", "renderer"],
        "draft.evaluate params",
      );
      const renderer = optionalString(request.params, "renderer");
      value = await (
        await openProject(manifestPath)
      ).evaluatePatch(
        asPatch(request.params.patch),
        renderer
          ? {
              renderer: renderer as
                | "deterministic"
                | "word"
                | "libreoffice"
                | "compare",
            }
          : {},
      );
      break;
    }
    case "draft.apply": {
      assertKeys(
        request.params,
        ["patch", "patchHash", "gate", "author", "message"],
        "draft.apply params",
      );
      const gate = optionalString(request.params, "gate");
      value = await (
        await openProject(manifestPath)
      ).applyPatch(asPatch(request.params.patch), {
        patchHash: requiredString(request.params, "patchHash"),
        ...(gate ? { gate: gate as "report" | "not-worse" | "pass" } : {}),
        author: actor(request.params.author),
        message: requiredString(request.params, "message"),
      });
      break;
    }
    case "review.add": {
      assertKeys(
        request.params,
        ["documentId", "revision", "blockId", "range", "author", "message"],
        "review.add params",
      );
      let range: AddReviewInput["range"];
      if (request.params.range !== undefined) {
        const raw = assertRecord(request.params.range, "range");
        assertKeys(raw, ["start", "end"], "range");
        if (!Number.isInteger(raw.start) || !Number.isInteger(raw.end))
          throw new AgentDocxError(
            "INVALID_ARGUMENT",
            "Review range must use integer offsets",
          );
        range = { start: raw.start as number, end: raw.end as number };
      }
      value = await (
        await openProject(manifestPath)
      ).addReview(requiredString(request.params, "documentId"), {
        revision: asRevision(request.params.revision, "revision") as
          | RevisionId
          | "HEAD",
        blockId: requiredString(
          request.params,
          "blockId",
        ) as AddReviewInput["blockId"],
        ...(range ? { range } : {}),
        author: actor(request.params.author),
        message: requiredString(request.params, "message"),
      });
      break;
    }
    case "review.resolve": {
      assertKeys(
        request.params,
        ["documentId", "revision", "annotationId", "author", "message"],
        "review.resolve params",
      );
      const input: ResolveReviewInput = {
        revision: asRevision(request.params.revision, "revision") as
          | RevisionId
          | "HEAD",
        annotationId: requiredString(
          request.params,
          "annotationId",
        ) as ResolveReviewInput["annotationId"],
        author: actor(request.params.author),
        message: requiredString(request.params, "message"),
      };
      value = await (
        await openProject(manifestPath)
      ).resolveReview(requiredString(request.params, "documentId"), input);
      break;
    }
    case "docx.export": {
      assertKeys(
        request.params,
        ["documentId", "revision", "mode", "baseRevision", "output", "options"],
        "docx.export params",
      );
      const mode = requiredString(request.params, "mode");
      const revision = asRevision(request.params.revision, "revision") as
        | RevisionId
        | "HEAD";
      const output = invocationPath(
        cwd,
        requiredString(request.params, "output"),
      );
      const options =
        request.params.options === undefined
          ? undefined
          : assertRecord(request.params.options, "options");
      let input: ExportDocxInput;
      if (mode === "clean" || mode === "pdf") {
        if (request.params.baseRevision !== undefined)
          throw new AgentDocxError(
            "INVALID_ARGUMENT",
            `${mode === "clean" ? "Clean" : "PDF"} DOCX export forbids baseRevision`,
          );
        input = {
          revision,
          mode,
          output,
          ...(options ? { options } : {}),
        } as ExportDocxInput;
      } else if (mode === "redline") {
        input = {
          revision,
          mode,
          baseRevision: asRevision(
            request.params.baseRevision,
            "baseRevision",
          ) as RevisionId | "HEAD",
          output,
          ...(options ? { options } : {}),
        } as ExportDocxInput;
      } else
        throw new AgentDocxError(
          "INVALID_ARGUMENT",
          "DOCX export mode must be clean or redline",
        );
      value = await (
        await openProject(manifestPath)
      ).exportDocx(requiredString(request.params, "documentId"), input);
      break;
    }
    case "docx.import": {
      assertKeys(
        request.params,
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
      const inspectOnly = request.params.inspectOnly;
      if (typeof inspectOnly !== "boolean")
        throw new AgentDocxError(
          "INVALID_ARGUMENT",
          "inspectOnly must be a boolean",
        );
      const inputPath = invocationPath(
        cwd,
        requiredString(request.params, "input"),
      );
      const attachmentPath = optionalString(request.params, "attachments");
      const attachments = attachmentPath
        ? { directory: invocationPath(cwd, attachmentPath) }
        : undefined;
      if (inspectOnly) {
        if (
          ["documentId", "output", "author", "message"].some(
            (key) => request.params[key] !== undefined,
          )
        )
          throw new AgentDocxError(
            "INVALID_ARGUMENT",
            "Inspect-only DOCX import is stateless",
          );
        value = await inspectDocx(
          inputPath,
          attachments ? { attachments } : {},
        );
      } else {
        const input: Extract<ImportDocxInput, { inspectOnly: false }> = {
          input: inputPath,
          inspectOnly: false,
          documentId: requiredString(request.params, "documentId"),
          output: invocationPath(cwd, requiredString(request.params, "output")),
          author: actor(request.params.author),
          message: requiredString(request.params, "message"),
          ...(attachments ? { attachments } : {}),
        };
        value = await (await openProject(manifestPath)).importDocx(input);
      }
      break;
    }
    case "docx.inspect": {
      assertKeys(request.params, ["input"], "docx.inspect params");
      value = await inspectDocxTemplate(
        await readFile(
          invocationPath(cwd, requiredString(request.params, "input")),
        ),
      );
      break;
    }
    case "docx.importRedline": {
      const project = await openProject(manifestPath);
      const attachmentPath = optionalString(request.params, "attachments");
      value = await project.importRedline({
        documentId: requiredDocumentId(request.params, "documentId"),
        input: invocationPath(cwd, requiredString(request.params, "input")),
        ...(attachmentPath
          ? { attachments: { directory: invocationPath(cwd, attachmentPath) } }
          : {}),
        author: actor(request.params.author),
        message: requiredString(request.params, "message"),
      });
      break;
    }
    case "filingSet.add": {
      const project = await openProject(manifestPath);
      value = await project.addFilingSet({
        id: requiredString(request.params, "id"),
        documentIds: request.params.documentIds as string[],
        ...(request.params.label === undefined
          ? {}
          : { label: request.params.label as string }),
        ...(request.params.pageCap === undefined
          ? {}
          : { pageCap: request.params.pageCap as number }),
      });
      break;
    }
    case "filingSet.remove": {
      const project = await openProject(manifestPath);
      value = await project.removeFilingSet(
        requiredString(request.params, "id"),
      );
      break;
    }
    case "filingSet.get": {
      const project = await openProject(manifestPath);
      value = await project.getFilingSet(requiredString(request.params, "id"));
      break;
    }
    case "filingSet.validate": {
      const project = await openProject(manifestPath);
      value = await project.validateFilingSet(
        requiredString(request.params, "id"),
      );
      break;
    }
  }
  const meta = responseMeta(request.action, request.params, value);
  return { request, project: projectDisplay, ...meta, value };
};

/**
 * Projects API values onto JSON-safe protocol values. Generated DOCX and
 * in-memory attachment payloads are intentionally omitted rather than encoded.
 */
export const serializeAgentValue = (
  value: unknown,
  cwd?: string,
): JsonValue => {
  const serialize = (
    candidate: unknown,
    key?: string,
  ): JsonValue | undefined => {
    if (candidate === undefined || candidate instanceof Uint8Array)
      return undefined;
    if (candidate === null) return null;
    if (typeof candidate === "string") {
      if (
        cwd &&
        key !== undefined &&
        ["manifestPath", "path", "storePath", "output"].includes(key) &&
        isAbsolute(candidate)
      )
        return publicPath(cwd, candidate);
      return candidate;
    }
    if (typeof candidate === "boolean") return candidate;
    if (typeof candidate === "number")
      return Number.isFinite(candidate) ? candidate : null;
    if (Array.isArray(candidate))
      return candidate.map((entry) => serialize(entry) ?? null);
    if (typeof candidate !== "object") return String(candidate);

    const output: Record<string, JsonValue> = {};
    for (const [childKey, child] of Object.entries(
      candidate as Record<string, unknown>,
    )) {
      if (childKey === "generatedDocx" || child instanceof Uint8Array) continue;
      if (
        childKey === "attachments" &&
        typeof child === "object" &&
        child !== null &&
        !Array.isArray(child) &&
        "files" in child
      )
        continue;
      const projected = serialize(child, childKey);
      if (projected !== undefined) output[childKey] = projected;
    }
    return output;
  };
  return serialize(value) ?? null;
};
