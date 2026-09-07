import type {
  DocumentConfigUpdate,
  ProjectDocumentInput,
} from "../project/contracts.js";
import { AgentDocxError } from "../types.js";
import {
  assertDate,
  assertKeys,
  assertRecord,
  filingKinds,
  manifestRelativePath,
  optionalBoolean,
  optionalEnum,
  optionalInteger,
  optionalRulePacks,
  optionalString,
  optionalText,
  profileIds,
  rendererModes,
  requiredDocumentId,
  requiredEnum,
  requiredInteger,
  requiredString,
  requiredText,
  rulePackIds,
} from "./protocol-primitives.js";

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
