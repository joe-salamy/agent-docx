import { AgentDocxError } from "../types.js";
import type {
  JsonValue,
  SourcePosition,
} from "../types.js";
import type {
  DeterministicResult,
  FilingKind,
} from "../measurement.js";
import type {
  LegalBlock,
  LegalDocument,
  RevisionId,
  RulePackId,
} from "./model.js";
import { visibleTextForBlock } from "./visible-text.js";

export type ValidationFinding = {
  checkId: string;
  status: "pass" | "fail" | "unknown";
  severity: "info" | "warning" | "error";
  source: "rule" | "package" | "template";
  message: string;
  citation?: string;
  sourceUrl?: string;
  effectiveDate?: string;
  positions?: readonly SourcePosition[];
  evidence: Readonly<Record<string, JsonValue>>;
  remediation?: Readonly<Record<string, JsonValue>>;
};

export type ValidationResult = {
  schemaVersion: 1;
  documentId: string;
  revision: RevisionId | null;
  rulePack: string | null;
  scope: {
    certification: false;
    checkedRuleIds: readonly string[];
    sourceSnapshots: readonly {
      id: string;
      sourceUrl: string;
      effectiveDate: string;
      sha256: `sha256:${string}`;
    }[];
    unmodeledProvisions: readonly string[];
  };
  status: "pass" | "fail" | "unknown";
  summary: { pass: number; fail: number; unknown: number };
  findings: readonly ValidationFinding[];
};

export type RuleCheckKind =
  | "length-alternative"
  | "page-size"
  | "margin-minimum"
  | "typeface"
  | "line-spacing"
  | "counted-lines-maximum"
  | "required-metadata"
  | "required-block"
  | "required-footer"
  | "reference-integrity";

export type UserRulePackLengthAlternative = {
  byFilingKind: Partial<
    Record<
      FilingKind | "default",
      {
        pages?: number;
        words?: number;
        monospacedLines?: number;
        complianceCertificateRequired?: boolean;
      }
    >
  >;
};
export type UserRulePackPageSize = {
  widthTwips: number;
  heightTwips: number;
};
export type UserRulePackMarginMinimum = { minimumTwips: number };
export type UserRulePackTypeface = {
  minimumTwips: number;
  mode: "proportional" | "monospaced";
  requireVerifiedPitch?: boolean;
};
export type UserRulePackLineSpacing = { doubleSpacedOrdinary?: boolean };
export type UserRulePackCountedLinesMaximum = { perPageMaximum: number };
export type UserRulePackRequiredMetadata = {
  fields: readonly string[];
  requireCounselComplete?: boolean;
  requireComplianceCertificate?: boolean;
};
export type UserRulePackRequiredBlock = { kinds: readonly string[] };
export type UserRulePackRequiredFooter = { requiredTokens: readonly string[] };
export type UserRulePackReferenceIntegrity = {};
export type RuleCheckParams =
  | UserRulePackLengthAlternative
  | UserRulePackPageSize
  | UserRulePackMarginMinimum
  | UserRulePackTypeface
  | UserRulePackLineSpacing
  | UserRulePackCountedLinesMaximum
  | UserRulePackRequiredMetadata
  | UserRulePackRequiredBlock
  | UserRulePackRequiredFooter
  | UserRulePackReferenceIntegrity;
export type UserRulePackCheck = {
  id: string;
  kind: RuleCheckKind;
  citation: string;
  predicate: string;
  params: RuleCheckParams;
};
export type UserRulePack = {
  id: string;
  sourceUrl: string;
  effectiveDate: string;
  sourceSha256: `sha256:${string}`;
  sourceExcerpt: string;
  checks: readonly UserRulePackCheck[];
  unmodeledProvisions: readonly string[];
};

export type BuiltInRuleCheck = {
  id: string;
  kind: RuleCheckKind;
  citation: string;
  predicate: string;
  params: RuleCheckParams;
};

export type BuiltInRulePack = {
  id: RulePackId;
  sourceUrl: string;
  effectiveDate: string;
  sourceSha256: `sha256:${string}`;
  sourceExcerpt: string;
  checks: readonly BuiltInRuleCheck[];
  unmodeledProvisions: readonly string[];
};

const frapSource = "frap-32-2024-12-01.txt";
const candSource = "cand-civil-2026-05-01.txt";

export const builtInRulePacks: Readonly<Record<RulePackId, BuiltInRulePack>> = {
  "frap-32@2024-12-01": {
    id: "frap-32@2024-12-01",
    sourceUrl:
      "https://www.uscourts.gov/forms-rules/current-rules-practice-procedure/federal-rules-appellate-procedure",
    effectiveDate: "2024-12-01",
    sourceSha256:
      "sha256:2dda704d3495539ff2e93c08029c058f0e12089949b2e6d02d94159f9f150786",
    sourceExcerpt: frapSource,
    checks: [
      {
        id: "frap32.length.principal",
        kind: "length-alternative",
        citation: "Fed. R. App. P. 32(a)(7)",
        predicate:
          "principal <= 30 pages OR <= 13000 words OR <= 1300 monospaced lines; non-page alternatives require a Rule 32(g) compliance certificate",
        params: {
          byFilingKind: {
            "principal-brief": {
              pages: 30,
              words: 13000,
              monospacedLines: 1300,
              complianceCertificateRequired: true,
            },
            },
        },
      },
      {
        id: "frap32.length.reply",
        kind: "length-alternative",
        citation: "Fed. R. App. P. 32(a)(7)",
        predicate:
          "reply <= 15 pages OR <= 6500 words OR <= 650 monospaced lines; non-page alternatives require a Rule 32(g) compliance certificate",
        params: {
          byFilingKind: {
            "reply-brief": {
              pages: 15,
              words: 6500,
              monospacedLines: 650,
              complianceCertificateRequired: true,
            },
            },
        },
      },
      {
        id: "frap32.page-size",
        kind: "page-size",
        citation: "Fed. R. App. P. 32(a)(4)",
        predicate: "page is exactly 8.5 by 11 inches",
        params: { widthTwips: 12240, heightTwips: 15840 },
      },
      {
        id: "frap32.margin",
        kind: "margin-minimum",
        citation: "Fed. R. App. P. 32(a)(4)",
        predicate: "each margin is at least one inch",
        params: { minimumTwips: 1440 },
      },
      {
        id: "frap32.typeface.proportional",
        kind: "typeface",
        citation: "Fed. R. App. P. 32(a)(5)",
        predicate: "proportional serif text is at least 14 point",
        params: { minimumTwips: 280, mode: "proportional" },
      },
      {
        id: "frap32.typeface.monospaced",
        kind: "typeface",
        citation: "Fed. R. App. P. 32(a)(5)",
        predicate:
          "monospaced text contains no more than 10.5 characters per inch",
        params: { minimumTwips: 280, mode: "monospaced" },
      },
      {
        id: "frap32.spacing",
        kind: "line-spacing",
        citation: "Fed. R. App. P. 32(a)(4)",
        predicate: "ordinary text is double spaced",
        params: { doubleSpacedOrdinary: true },
      },
    ],
    unmodeledProvisions: [
      "Rule 32(a)(1)-(3), (6), and formatting conditions not represented by the legal IR",
      "Court-specific local-rule exclusions beyond an explicitly cited directive",
    ],
  },
  "cand-civil@2026-05-01": {
    id: "cand-civil@2026-05-01",
    sourceUrl:
      "https://cand.uscourts.gov/rules-forms-fees/local-rules/civil-local-rules",
    effectiveDate: "2026-05-01",
    sourceSha256:
      "sha256:82383622a53a6ffe15bdc7931c263db4164c2e4a263a79fe4e53eb80d86dcbb5",
    sourceExcerpt: candSource,
    checks: [
      {
        id: "cand.length.motion",
        kind: "length-alternative",
        citation: "Civil L.R. 7-2(b)",
        predicate: "motion document <= 25 pages",
        params: {
          byFilingKind: { "motion-document": { pages: 25 } },
        },
      },
      {
        id: "cand.length.opposition",
        kind: "length-alternative",
        citation: "Civil L.R. 7-3(a), 7-3(c)",
        predicate: "opposition text <= 25 pages",
        params: {
          byFilingKind: { "opposition-text": { pages: 25 } },
        },
      },
      {
        id: "cand.length.reply",
        kind: "length-alternative",
        citation: "Civil L.R. 7-4(b)",
        predicate: "reply text <= 15 pages",
        params: {
          byFilingKind: { default: { pages: 15 }, "reply-text": { pages: 15 } },
        },
      },
      {
        id: "cand.lines",
        kind: "counted-lines-maximum",
        citation: "Civil L.R. 3-4(c)(2)",
        predicate: "each page has at most 28 counted lines",
        params: { perPageMaximum: 28 },
      },
      {
        id: "cand.typeface",
        kind: "typeface",
        citation: "Civil L.R. 3-4(c)(2)",
        predicate:
          "all text is at least 12 point in a verified proportional serif face",
        params: {
          minimumTwips: 240,
          mode: "proportional",
          requireVerifiedPitch: true,
        },
      },
      {
        id: "cand.spacing",
        kind: "line-spacing",
        citation: "Civil L.R. 3-4(c)(2)",
        predicate: "ordinary text is double spaced",
        params: { doubleSpacedOrdinary: true },
      },
      {
        id: "cand.first-page",
        kind: "required-metadata",
        citation: "Civil L.R. 3-4(a)",
        predicate: "counsel, court, case, and document-title facts are present",
        params: {
          fields: [
            "court",
            "jurisdiction",
            "caseName",
            "docketNumber",
            "documentTitle",
          ],
          requireCounselComplete: true,
        },
      },
      {
        id: "cand.footer",
        kind: "required-footer",
        citation: "Civil L.R. 3-4(c)(3)",
        predicate: "footer contains case title and case number",
        params: {
          requiredTokens: ["{{caseName}}", "{{docketNumber}}"],
        },
      },
    ],
    unmodeledProvisions: [
      "Judge-specific and conditional legends",
      "Individual judge standing orders and page-limit exceptions",
    ],
  },
};
const userRulePackKinds: readonly RuleCheckKind[] = [
  "length-alternative",
  "page-size",
  "margin-minimum",
  "typeface",
  "line-spacing",
  "counted-lines-maximum",
  "required-metadata",
  "required-block",
  "required-footer",
  "reference-integrity",
];
const metadataStringFields = new Set([
  "court",
  "jurisdiction",
  "caseName",
  "docketNumber",
  "documentTitle",
  "filingDate",
]);
const legalBlockKinds = new Set([
  "paragraph",
  "blockquote",
  "heading",
  "numbered-paragraph",
  "list",
  "table",
  "caption",
  "toc",
  "toa",
  "signature",
  "certificate",
  "exhibit",
  "length-exclusion",
  "image",
  "pagebreak",
  "thematic-break",
  "sectionbreak",
]);
const isRecord = (value: unknown): value is Record<string, unknown> =>
  value !== null && typeof value === "object" && !Array.isArray(value);
const userPackError = (label: string, message: string): never => {
  throw new AgentDocxError("RULE_PACK_INVALID", `${label}: ${message}`);
};
const exactKeys = (
  value: Record<string, unknown>,
  keys: readonly string[],
  label: string,
): void => {
  const allowed = new Set(keys);
  if (Object.keys(value).some((key) => !allowed.has(key)))
    userPackError(label, "contains unknown properties");
};
const requiredKeys = (
  value: Record<string, unknown>,
  keys: readonly string[],
  label: string,
): void => {
  if (keys.some((key) => !(key in value)))
    userPackError(label, "is missing required properties");
};
const positiveInteger = (value: unknown): value is number =>
  typeof value === "number" && Number.isInteger(value) && value > 0;
const nonNegativeInteger = (value: unknown): value is number =>
  typeof value === "number" && Number.isInteger(value) && value >= 0;
const stringArray = (
  value: unknown,
  label: string,
  requireNonEmpty = true,
): readonly string[] => {
  if (
    !Array.isArray(value) ||
    (requireNonEmpty && value.length === 0) ||
    value.some((entry) => typeof entry !== "string" || entry.length === 0)
  )
    userPackError(
      label,
      requireNonEmpty
        ? "must be a non-empty string array"
        : "must be a string array",
    );
  return value as string[];
};
const validateRuleCheckParams = (
  kind: RuleCheckKind,
  value: unknown,
  label: string,
): RuleCheckParams => {
  if (!isRecord(value)) return userPackError(label, "params must be an object");
  switch (kind) {
    case "length-alternative": {
      exactKeys(value, ["byFilingKind"], label);
      requiredKeys(value, ["byFilingKind"], label);
      if (!isRecord(value.byFilingKind))
        return userPackError(label, "byFilingKind must be an object");
      const alternatives: Record<string, unknown> = {};
      for (const [filingKind, raw] of Object.entries(value.byFilingKind)) {
        if (
          filingKind !== "default" &&
          ![
            "principal-brief",
            "reply-brief",
            "motion-document",
            "opposition-text",
            "reply-text",
          ].includes(filingKind)
        )
          userPackError(label, `unknown filing kind: ${filingKind}`);
        if (!isRecord(raw))
          return userPackError(label, `${filingKind} must be an object`);
        exactKeys(
          raw,
          [
            "pages",
            "words",
            "monospacedLines",
            "complianceCertificateRequired",
          ],
          `${label}.${filingKind}`,
        );
        for (const key of ["pages", "words", "monospacedLines"] as const)
          if (raw[key] !== undefined && !positiveInteger(raw[key]))
            userPackError(
              `${label}.${filingKind}`,
              `${key} must be a positive integer`,
            );
        if (
          raw.complianceCertificateRequired !== undefined &&
          typeof raw.complianceCertificateRequired !== "boolean"
        )
          userPackError(
            `${label}.${filingKind}`,
            "complianceCertificateRequired must be boolean",
          );
        alternatives[filingKind] = raw;
      }
      if (Object.keys(alternatives).length === 0)
        userPackError(label, "byFilingKind must contain an entry");
      return { byFilingKind: alternatives } as RuleCheckParams;
    }
    case "page-size":
      exactKeys(value, ["widthTwips", "heightTwips"], label);
      requiredKeys(value, ["widthTwips", "heightTwips"], label);
      if (
        !positiveInteger(value.widthTwips) ||
        !positiveInteger(value.heightTwips)
      )
        userPackError(label, "page dimensions must be positive integers");
      return value as RuleCheckParams;
    case "margin-minimum":
      exactKeys(value, ["minimumTwips"], label);
      requiredKeys(value, ["minimumTwips"], label);
      if (!nonNegativeInteger(value.minimumTwips))
        userPackError(label, "minimumTwips must be a non-negative integer");
      return value as RuleCheckParams;
    case "typeface":
      exactKeys(value, ["minimumTwips", "mode", "requireVerifiedPitch"], label);
      requiredKeys(value, ["minimumTwips", "mode"], label);
      if (!nonNegativeInteger(value.minimumTwips))
        userPackError(label, "minimumTwips must be a non-negative integer");
      if (value.mode !== "proportional" && value.mode !== "monospaced")
        userPackError(label, "mode is invalid");
      if (
        value.requireVerifiedPitch !== undefined &&
        typeof value.requireVerifiedPitch !== "boolean"
      )
        userPackError(label, "requireVerifiedPitch must be boolean");
      return value as RuleCheckParams;
    case "line-spacing":
      exactKeys(value, ["doubleSpacedOrdinary"], label);
      if (
        value.doubleSpacedOrdinary !== undefined &&
        typeof value.doubleSpacedOrdinary !== "boolean"
      )
        userPackError(label, "doubleSpacedOrdinary must be boolean");
      return value as RuleCheckParams;
    case "counted-lines-maximum":
      exactKeys(value, ["perPageMaximum"], label);
      requiredKeys(value, ["perPageMaximum"], label);
      if (!positiveInteger(value.perPageMaximum))
        userPackError(label, "perPageMaximum must be a positive integer");
      return value as RuleCheckParams;
    case "required-metadata": {
      exactKeys(
        value,
        ["fields", "requireCounselComplete", "requireComplianceCertificate"],
        label,
      );
      requiredKeys(value, ["fields"], label);
      const fields = stringArray(value.fields, `${label}.fields`, false);
      for (const field of fields)
        if (!metadataStringFields.has(field))
          userPackError(`${label}.fields`, `unknown metadata field: ${field}`);
      for (const key of [
        "requireCounselComplete",
        "requireComplianceCertificate",
      ] as const)
        if (value[key] !== undefined && typeof value[key] !== "boolean")
          userPackError(label, `${key} must be boolean`);
      return { ...value, fields } as RuleCheckParams;
    }
    case "required-block": {
      exactKeys(value, ["kinds"], label);
      requiredKeys(value, ["kinds"], label);
      const kinds = stringArray(value.kinds, `${label}.kinds`, false);
      for (const blockKind of kinds)
        if (!legalBlockKinds.has(blockKind))
          userPackError(
            `${label}.kinds`,
            `unknown legal block kind: ${blockKind}`,
          );
      return { kinds } as RuleCheckParams;
    }
    case "required-footer":
      exactKeys(value, ["requiredTokens"], label);
      requiredKeys(value, ["requiredTokens"], label);
      return {
        requiredTokens: stringArray(
          value.requiredTokens,
          `${label}.requiredTokens`,
          false,
        ),
      } as RuleCheckParams;
    case "reference-integrity":
      exactKeys(value, [], label);
      return {} as RuleCheckParams;
  }
};

export const validateUserRulePack = (
  value: unknown,
  label: string,
): UserRulePack => {
  if (!isRecord(value)) return userPackError(label, "pack must be an object");
  const pack = value as Record<string, unknown>;
  exactKeys(
    pack,
    [
      "id",
      "sourceUrl",
      "effectiveDate",
      "sourceSha256",
      "sourceExcerpt",
      "checks",
      "unmodeledProvisions",
    ],
    label,
  );
  if (
    typeof pack.id !== "string" ||
    !/^[a-z0-9][a-z0-9@.-]{0,127}$/.test(pack.id)
  )
    userPackError(label, "id is invalid");
  if (typeof pack.sourceUrl !== "string" || !/^https:\/\//.test(pack.sourceUrl))
    userPackError(label, "sourceUrl is invalid");
  if (
    typeof pack.effectiveDate !== "string" ||
    !/^\d{4}-\d{2}-\d{2}$/.test(pack.effectiveDate)
  )
    userPackError(label, "effectiveDate is invalid");
  if (
    typeof pack.sourceSha256 !== "string" ||
    !/^sha256:[a-f0-9]{64}$/.test(pack.sourceSha256)
  )
    userPackError(label, "sourceSha256 is invalid");
  if (typeof pack.sourceExcerpt !== "string" || pack.sourceExcerpt.length === 0)
    userPackError(label, "sourceExcerpt is invalid");
  const id = pack.id as string;
  const sourceUrl = pack.sourceUrl as string;
  const effectiveDate = pack.effectiveDate as string;
  const sourceSha256 = pack.sourceSha256 as `sha256:${string}`;
  const sourceExcerpt = pack.sourceExcerpt as string;
  if (!Array.isArray(pack.checks) || pack.checks.length === 0)
    userPackError(label, "checks must be a non-empty array");
  const ids = new Set<string>();
  const checks = (pack.checks as unknown[]).map(
    (raw: unknown, index: number): UserRulePack["checks"][number] => {
      const checkLabel = `${label}.checks[${index}]`;
      if (!isRecord(raw))
        return userPackError(checkLabel, "check must be an object");
      const check = raw as Record<string, unknown>;
      exactKeys(
        check,
        ["id", "kind", "citation", "predicate", "params"],
        checkLabel,
      );
      if (typeof check.id !== "string" || !/^[a-z0-9.-]+$/.test(check.id))
        userPackError(checkLabel, "id is invalid");
      const checkId = check.id as string;
      if (ids.has(checkId))
        userPackError(checkLabel, `duplicate check id: ${checkId}`);
      ids.add(checkId);
      const kindValue = check.kind as string;
      if (
        typeof check.kind !== "string" ||
        !(userRulePackKinds as readonly string[]).includes(kindValue)
      )
        userPackError(checkLabel, "kind is invalid");
      if (typeof check.citation !== "string" || check.citation.length === 0)
        userPackError(checkLabel, "citation is invalid");
      if (typeof check.predicate !== "string" || check.predicate.length === 0)
        userPackError(checkLabel, "predicate is invalid");
      const citation = check.citation as string;
      const predicate = check.predicate as string;
      const kind = check.kind as RuleCheckKind;
      return {
        id: checkId,
        kind,
        citation,
        predicate,
        params: validateRuleCheckParams(
          kind,
          check.params,
          `${checkLabel}.params`,
        ),
      };
    },
  );
  const unmodeledProvisions = stringArray(
    pack.unmodeledProvisions,
    `${label}.unmodeledProvisions`,
    false,
  );
  return {
    id,
    sourceUrl,
    effectiveDate,
    sourceSha256,
    sourceExcerpt,
    checks,
    unmodeledProvisions,
  };
};

type RuleMeasurement = {
  pageCount: number;
  deterministic: Pick<
    DeterministicResult,
    | "pageCount"
    | "totalVisualLines"
    | "visualLinesByPage"
    | "countedLinesByPage"
    | "profile"
  >;
};

export type ValidationInput = {
  revision?: RevisionId | null;
  rulePack?: RulePackId;
  customPacks?: readonly UserRulePack[];
  filingKind?: FilingKind;
  measurement?: RuleMeasurement;
};
type CountedEntry = {
  id: LegalBlock["id"];
  kind: string;
  text: string;
  position: SourcePosition;
};

const textFor = (block: LegalBlock): string => visibleTextForBlock(block);

const flattenBlocks = (blocks: readonly LegalBlock[]): LegalBlock[] => {
  const flattened: LegalBlock[] = [];
  for (const block of blocks) {
    flattened.push(block);
    if (block.kind === "exhibit" || block.kind === "length-exclusion")
      flattened.push(...flattenBlocks(block.blocks));
  }
  return flattened;
};

const countedEntries = (document: LegalDocument): readonly CountedEntry[] => {
  const entries: CountedEntry[] = [];
  const excluded = new Set([
    "caption",
    "toc",
    "toa",
    "signature",
    "certificate",
    "pagebreak",
    "sectionbreak",
    "thematic-break",
  ]);
  const visit = (
    blocks: readonly LegalBlock[],
    inheritedExclusion = false,
  ): void => {
    for (const block of blocks) {
      const excludedHere =
        inheritedExclusion || block.kind === "length-exclusion";
      if (!excludedHere && !excluded.has(block.kind))
        entries.push({
          id: block.id,
          kind: block.kind,
          text: textFor(block),
          position: block.position,
        });
      if (block.kind === "exhibit") visit(block.blocks, excludedHere);
      if (block.kind === "length-exclusion") visit(block.blocks, true);
    }
  };
  visit(document.blocks);
  for (const footnote of document.footnotes)
    entries.push({
      id: footnote.id,
      kind: "footnote",
      text: footnote.paragraphs
        .flatMap((paragraph) => paragraph.runs)
        .map((run) => run.text)
        .join(" "),
      position: footnote.position,
    });
  return entries;
};

const wordsIn = (entries: readonly CountedEntry[]): number =>
  entries
    .map((entry) => entry.text)
    .join(" ")
    .match(/[\p{L}\p{N}]+/gu)?.length ?? 0;

const doubleSpaced = (
  spacing:
    | { rule: "auto"; numerator: number; denominator: 240 }
    | { rule: "exact" | "atLeast"; twips: number },
): boolean =>
  spacing.rule === "auto" ? spacing.numerator >= 480 : spacing.twips >= 480;

type RulePackForFinding = {
  id: string;
  sourceUrl: string;
  effectiveDate: string;
  sourceSha256: `sha256:${string}`;
  checks: readonly { id: string; citation: string; predicate: string }[];
};
const ruleFinding = (
  pack: RulePackForFinding,
  checkId: string,
  status: ValidationFinding["status"],
  message: string,
  evidence: Readonly<Record<string, JsonValue>>,
  positions?: readonly SourcePosition[],
): ValidationFinding => ({
  checkId,
  status,
  severity: "error",
  source: "rule",
  message,
  ...(pack.checks.find((check) => check.id === checkId)?.citation
    ? {
        citation: pack.checks.find((check) => check.id === checkId)!.citation,
      }
    : {}),
  sourceUrl: pack.sourceUrl,
  effectiveDate: pack.effectiveDate,
  ...(positions && positions.length > 0 ? { positions } : {}),
  evidence: {
    sourceSnapshotId: pack.id,
    sourceSha256: pack.sourceSha256,
    predicate:
      pack.checks.find((check) => check.id === checkId)?.predicate ?? null,
    ...evidence,
  },
});

const packageFinding = (
  checkId: string,
  status: ValidationFinding["status"],
  severity: ValidationFinding["severity"],
  message: string,
  evidence: Readonly<Record<string, JsonValue>>,
): ValidationFinding => ({
  checkId,
  status,
  severity,
  source: "package",
  message,
  evidence,
});

const metadataFinding = (document: LegalDocument): ValidationFinding => {
  const required: [string, string][] = [
    ["court", document.metadata.court],
    ["jurisdiction", document.metadata.jurisdiction],
    ["caseName", document.metadata.caseName],
    ["docketNumber", document.metadata.docketNumber],
    ["documentTitle", document.metadata.documentTitle],
  ];
  const missing = required
    .filter(([, value]) => value.trim().length === 0)
    .map(([key]) => key);
  return packageFinding(
    "package.metadata",
    missing.length === 0 ? "pass" : "fail",
    "warning",
    missing.length === 0
      ? "Core matter metadata is present"
      : `Missing core metadata: ${missing.join(", ")}`,
    { missing },
  );
};

const runsForBlock = (block: LegalBlock) => {
  if ("runs" in block) return block.runs;
  if (block.kind === "list")
    return block.items.flatMap((item) =>
      item.paragraphs.flatMap((paragraph) => paragraph.runs),
    );
  if (block.kind === "table")
    return block.rows.flatMap((row) =>
      row.flatMap((cell) =>
        cell.paragraphs.flatMap((paragraph) => paragraph.runs),
      ),
    );
  return [];
};

const duplicates = (values: readonly string[]): readonly string[] =>
  values.filter((value, index) => values.indexOf(value) !== index);

const referenceFinding = (document: LegalDocument): ValidationFinding => {
  const blocks = flattenBlocks(document.blocks);
  const blockIds = new Set(blocks.map((block) => block.id));
  const duplicateBlockIds = duplicates(blocks.map((block) => block.id));
  const duplicateFootnotes = duplicates(
    document.footnotes.map((footnote) => footnote.label),
  );
  const duplicateCounsel = duplicates(
    document.metadata.counsel.map((entry) => entry.id),
  );
  const duplicateCertificates = duplicates(
    document.metadata.certificates.map((entry) => entry.id),
  );
  const duplicateAnnotations = duplicates(
    document.annotations.map((annotation) => annotation.id),
  );
  const counsel = new Set(document.metadata.counsel.map((entry) => entry.id));
  const certificates = new Set(
    document.metadata.certificates.map((entry) => entry.id),
  );
  const footnotes = new Set(
    document.footnotes.map((footnote) => footnote.label),
  );
  const unresolvedBlocks = blocks.flatMap((block) => {
    if (block.kind === "signature" && !counsel.has(block.counselId))
      return [block.id];
    if (block.kind === "certificate" && !certificates.has(block.certificateId))
      return [block.id];
    return [];
  });
  const runs = [
    ...blocks.flatMap(runsForBlock),
    ...document.footnotes.flatMap((footnote) =>
      footnote.paragraphs.flatMap((paragraph) => paragraph.runs),
    ),
  ];
  const unresolvedFootnotes = runs
    .flatMap((run) => (run.footnoteId ? [run.footnoteId] : []))
    .filter((id) => !footnotes.has(id));
  const unresolvedTargets = runs
    .flatMap((run) => (run.referenceTarget ? [run.referenceTarget] : []))
    .filter((id) => !blockIds.has(id));
  const invalidAnnotations = document.annotations
    .filter((annotation) => !blockIds.has(annotation.blockId))
    .map((annotation) => annotation.id);
  const valid =
    duplicateBlockIds.length === 0 &&
    duplicateFootnotes.length === 0 &&
    duplicateCounsel.length === 0 &&
    duplicateCertificates.length === 0 &&
    duplicateAnnotations.length === 0 &&
    unresolvedBlocks.length === 0 &&
    unresolvedFootnotes.length === 0 &&
    unresolvedTargets.length === 0 &&
    invalidAnnotations.length === 0;
  return packageFinding(
    "package.references",
    valid ? "pass" : "fail",
    "error",
    valid
      ? "Block, footnote, counsel, certificate, and cross references are unique and resolved"
      : "Document contains duplicate or unresolved references",
    {
      duplicateBlockIds,
      duplicateFootnotes,
      duplicateCounsel,
      duplicateCertificates,
      duplicateAnnotations,
      unresolvedBlockIds: unresolvedBlocks,
      unresolvedFootnotes,
      unresolvedTargets,
      invalidAnnotations,
    },
  );
};

const structureFinding = (document: LegalDocument): ValidationFinding => {
  const kinds = new Set(
    flattenBlocks(document.blocks).map((block) => block.kind),
  );
  const required = ["caption", "signature", "certificate"] as const;
  const missing = required.filter((kind) => !kinds.has(kind));
  return packageFinding(
    "package.structure",
    missing.length === 0 ? "pass" : "fail",
    "warning",
    missing.length === 0
      ? "Expected filing structure is present"
      : `Product-convention blocks are absent: ${missing.join(", ")}`,
    { missing },
  );
};

const typefaceStatus = (
  measurement: ValidationInput["measurement"],
  minimumTwips: number,
  mode: "proportional" | "monospaced",
  requireVerifiedPitch = false,
): {
  status: ValidationFinding["status"];
  evidence: Readonly<Record<string, JsonValue>>;
} => {
  if (!measurement)
    return {
      status: "unknown",
      evidence: { reason: "No deterministic measurement was supplied", mode },
    };
  const profile = measurement.deterministic.profile;
  const family = profile.requestedFontFamily;
  const monospaced = /mono|courier|consolas|menlo/i.test(family);
  const serif = /serif|times|georgia|garamond|bookman|cambria/i.test(family);
  const sizes = [
    profile.body.fontSizeTwips,
    profile.blockquote.fontSizeTwips,
    profile.list.fontSizeTwips,
    profile.footnote.fontSizeTwips,
    profile.table.body.fontSizeTwips,
    profile.table.header.fontSizeTwips,
    ...Object.values(profile.headings).map((style) => style.fontSizeTwips),
  ];
  const sizePasses = sizes.every((size) => size >= minimumTwips);
  const evidence = {
    family,
    classification: monospaced
      ? "monospaced"
      : serif
        ? "proportional-serif"
        : "unknown",
    sizes,
    minimumTwips,
    charactersPerInch: profile.maxCharactersPerInch,
  };
  if (mode === "monospaced") {
    if (!monospaced)
      return {
        status: "unknown",
        evidence: { ...evidence, reason: "Font is not verified monospaced" },
      };
    if (profile.maxCharactersPerInch === null)
      return {
        status: "unknown",
        evidence: { ...evidence, reason: "Characters-per-inch is unproved" },
      };
    return {
      status: profile.maxCharactersPerInch <= 10.5 ? "pass" : "fail",
      evidence,
    };
  }
  if (monospaced || !serif)
    return {
      status: "unknown",
      evidence: {
        ...evidence,
        reason: "Proportional serif classification is unproved",
      },
    };
  if (!sizePasses) return { status: "fail", evidence };
  if (requireVerifiedPitch)
    return {
      status: "unknown",
      evidence: {
        ...evidence,
        pitchVerified: false,
        manualVerificationRequired: true,
        reason:
          "The selected font's applicable pitch is not independently verified",
      },
    };
  return { status: "pass", evidence };
};

const spacingStatus = (
  measurement: ValidationInput["measurement"],
): {
  status: ValidationFinding["status"];
  evidence: Readonly<Record<string, JsonValue>>;
} => {
  if (!measurement)
    return {
      status: "unknown",
      evidence: { reason: "No deterministic measurement was supplied" },
    };
  const profile = measurement.deterministic.profile;
  const ordinary = [profile.body.lineSpacing, profile.list.lineSpacing];
  return {
    status: ordinary.every(doubleSpaced) ? "pass" : "fail",
    evidence: { ordinaryDoubleSpaced: ordinary.map(doubleSpaced) },
  };
};

const builtInCheck = (
  pack: BuiltInRulePack,
  kind: RuleCheckKind,
  select: (params: RuleCheckParams) => boolean = () => true,
): BuiltInRuleCheck => {
  const check = pack.checks.find(
    (candidate) => candidate.kind === kind && select(candidate.params),
  );
  if (!check)
    throw new AgentDocxError(
      "INTERNAL_ERROR",
      `Built-in rule pack ${pack.id} is missing a ${kind} check`,
    );
  return check;
};


const applyFrapChecks = (
  document: LegalDocument,
  input: ValidationInput,
  pack: BuiltInRulePack,
): ValidationFinding[] => {
  const measurement = input.measurement;
  const profile = measurement?.deterministic.profile;
  const counted = countedEntries(document);
  const words = wordsIn(counted);
  const lines = measurement?.deterministic.totalVisualLines ?? null;
  const pages = measurement?.pageCount ?? null;
  const filingKind =
    input.filingKind === "reply-brief" ? "reply-brief" : "principal-brief";
  const lengthCheck = builtInCheck(
    pack,
    "length-alternative",
    (params) =>
      "byFilingKind" in params &&
      params.byFilingKind[filingKind] !== undefined,
  );
  const configuredLimits = (
    lengthCheck.params as UserRulePackLengthAlternative
  ).byFilingKind[filingKind];
  if (!configuredLimits)
    throw new AgentDocxError(
      "INTERNAL_ERROR",
      `Built-in rule pack ${pack.id} has no limits for ${filingKind}`,
    );
  const limits = {
    pages: configuredLimits.pages!,
    words: configuredLimits.words!,
    lines: configuredLimits.monospacedLines!,
  };
  const requiresComplianceCertificate =
    configuredLimits.complianceCertificateRequired === true;
  const compliance = document.metadata.certificates.find(
    (certificate) => certificate.kind === "compliance",
  );
  const complianceBlock = compliance
    ? flattenBlocks(document.blocks).find(
        (block) =>
          block.kind === "certificate" && block.certificateId === compliance.id,
      )
    : undefined;
  const alternatives = [
    {
      name: "pages",
      status:
        pages === null ? "unknown" : pages <= limits.pages ? "pass" : "fail",
    },
    {
      name: "words",
      status:
        requiresComplianceCertificate && complianceBlock === undefined
          ? "fail"
          : words <= limits.words
            ? "pass"
            : "fail",
    },
    {
      name: "monospaced-lines",
      status:
        lines === null
          ? "unknown"
          : requiresComplianceCertificate && complianceBlock === undefined
            ? "fail"
            : lines <= limits.lines
              ? "pass"
              : "fail",
    },
  ] as const;
  const lengthStatus = alternatives.some(
    (alternative) => alternative.status === "pass",
  )
    ? "pass"
    : alternatives.every((alternative) => alternative.status === "fail")
      ? "fail"
      : "unknown";
  const pageSizeCheck = builtInCheck(pack, "page-size");
  const pageSize = pageSizeCheck.params as UserRulePackPageSize;
  const sizePasses =
    profile?.page.widthTwips === pageSize.widthTwips &&
    profile.page.heightTwips === pageSize.heightTwips;
  const marginCheck = builtInCheck(pack, "margin-minimum");
  const minimumMargin =
    marginCheck.params as UserRulePackMarginMinimum;
  const margins = profile?.page.marginsTwips;
  const marginPasses = margins
    ? Object.values(margins).every(
        (margin) => margin >= minimumMargin.minimumTwips,
      )
    : false;
  const monospaced = /mono|courier|consolas|menlo/i.test(
    profile?.requestedFontFamily ?? "",
  );
  const typefaceCheck = builtInCheck(
    pack,
    "typeface",
    (params) =>
      "mode" in params &&
      params.mode === (monospaced ? "monospaced" : "proportional"),
  );
  const typefaceParameters =
    typefaceCheck.params as UserRulePackTypeface;
  const typeface = typefaceStatus(
    measurement,
    typefaceParameters.minimumTwips,
    typefaceParameters.mode,
    typefaceParameters.requireVerifiedPitch ?? false,
  );
  const spacingCheck = builtInCheck(pack, "line-spacing");
  const spacing = spacingStatus(measurement);
  return [
    ruleFinding(
      pack,
      lengthCheck.id,
      lengthStatus,
      lengthStatus === "pass"
        ? "A Rule 32(a)(7) length alternative is satisfied"
        : "No Rule 32(a)(7) length alternative is proved",
      {
        filingKind: input.filingKind ?? "principal-brief",
        pages,
        words,
        monospacedLines: lines,
        limits,
        alternatives,
        complianceCertificateId: compliance?.id ?? null,
        complianceBlockId: complianceBlock?.id ?? null,
        countedBlockIds: counted.map((entry) => entry.id),
      },
      counted.map((entry) => entry.position),
    ),
    ruleFinding(
      pack,
      pageSizeCheck.id,
      profile ? (sizePasses ? "pass" : "fail") : "unknown",
      sizePasses
        ? "Page size is 8.5 by 11 inches"
        : "Page size is not verified as 8.5 by 11 inches",
      {
        widthTwips: profile?.page.widthTwips ?? null,
        heightTwips: profile?.page.heightTwips ?? null,
      },
    ),
    ruleFinding(
      pack,
      marginCheck.id,
      profile ? (marginPasses ? "pass" : "fail") : "unknown",
      marginPasses
        ? "All page margins are at least one inch"
        : "One or more page margins are below one inch",
      {
        marginsTwips: margins ?? null,
        minimumTwips: minimumMargin.minimumTwips,
      },
    ),
    ruleFinding(
      pack,
      typefaceCheck.id,
      typeface.status,
      typeface.status === "pass"
        ? "Typeface evidence satisfies the selected Rule 32 check"
        : "Typeface evidence does not prove the selected Rule 32 check",
      typeface.evidence,
    ),
    ruleFinding(
      pack,
      spacingCheck.id,
      spacing.status,
      spacing.status === "pass"
        ? "Ordinary text is double spaced"
        : "Ordinary text spacing does not satisfy Rule 32",
      spacing.evidence,
    ),
  ];
};

const applyCandChecks = (
  document: LegalDocument,
  input: ValidationInput,
  pack: BuiltInRulePack,
): ValidationFinding[] => {
  const measurement = input.measurement;
  const pages = measurement?.pageCount ?? null;
  const kind = input.filingKind;
  const selectedLengthKind =
    kind === "motion-document"
      ? "motion-document"
      : kind === "opposition-text"
        ? "opposition-text"
        : "reply-text";
  const lengthCheck = builtInCheck(
    pack,
    "length-alternative",
    (params) =>
      "byFilingKind" in params &&
      params.byFilingKind[selectedLengthKind] !== undefined,
  );
  const lengthParameters = (
    lengthCheck.params as UserRulePackLengthAlternative
  ).byFilingKind[selectedLengthKind];
  if (!lengthParameters?.pages)
    throw new AgentDocxError(
      "INTERNAL_ERROR",
      `Built-in rule pack ${pack.id} has no page limit for ${selectedLengthKind}`,
    );
  const typefaceCheck = builtInCheck(pack, "typeface");
  const typefaceParameters =
    typefaceCheck.params as UserRulePackTypeface;
  const typeface = typefaceStatus(
    measurement,
    typefaceParameters.minimumTwips,
    typefaceParameters.mode,
    typefaceParameters.requireVerifiedPitch ?? false,
  );
  const spacingCheck = builtInCheck(pack, "line-spacing");
  const spacing = spacingStatus(measurement);
  const countedLines = measurement?.deterministic.countedLinesByPage;
  const visualLines = measurement?.deterministic.visualLinesByPage;
  const lines = countedLines ?? visualLines ?? null;
  const linesCheck = builtInCheck(pack, "counted-lines-maximum");
  const linesParameters =
    linesCheck.params as UserRulePackCountedLinesMaximum;
  const linesStatus =
    lines === null
      ? "unknown"
      : lines.every(
            (lineCount) => lineCount <= linesParameters.perPageMaximum,
          )
        ? "pass"
        : "fail";
  const metadataCheck = builtInCheck(pack, "required-metadata");
  const metadataParameters =
    metadataCheck.params as UserRulePackRequiredMetadata;
  const missingMetadata = metadataParameters.fields.flatMap((field) => {
    const value =
      document.metadata[field as keyof typeof document.metadata];
    return typeof value === "string" && value.trim().length > 0 ? [] : [field];
  });
  const incompleteCounsel = document.metadata.counsel
    .filter(
      (counsel) =>
        !counsel.name.trim() ||
        !counsel.barNumber?.trim() ||
        !counsel.addressLines?.some((line) => line.trim()) ||
        !counsel.phone?.trim() ||
        !counsel.email?.trim(),
    )
    .map((counsel) => counsel.id);
  const footerCheck = builtInCheck(pack, "required-footer");
  const footerParameters =
    footerCheck.params as UserRulePackRequiredFooter;
  const footer = document.chrome.footers?.default ?? "";
  const missingFooterTokens = footerParameters.requiredTokens.filter(
    (token) => !footer.includes(token),
  );
  const footerPasses = missingFooterTokens.length === 0;
  const metadataPasses =
    missingMetadata.length === 0 &&
    (!metadataParameters.requireCounselComplete ||
      document.metadata.counsel.length > 0) &&
    (!metadataParameters.requireCounselComplete ||
      incompleteCounsel.length === 0);
  return [
    ruleFinding(
      pack,
      lengthCheck.id,
      pages === null
        ? "unknown"
        : pages <= lengthParameters.pages
          ? "pass"
          : "fail",
      pages !== null && pages <= lengthParameters.pages
        ? "Document page count is within the selected Civil Local Rule limit"
        : "Document page count exceeds or cannot prove the selected Civil Local Rule limit",
      {
        filingKind: kind ?? "reply-text",
        totalPages: pages,
        pagesOfBriefText: pages,
        limit: lengthParameters.pages,
      },
    ),
    ruleFinding(
      pack,
      linesCheck.id,
      linesStatus,
      linesStatus === "pass"
        ? "Each page has at most 28 counted lines"
        : "One or more pages exceed or cannot prove the 28-line limit",
      {
        lines: lines ?? [],
        countedLineSource: countedLines
          ? "deterministic-counted-lines"
          : "deterministic-visual-lines",
        maximum: linesParameters.perPageMaximum,
      },
    ),
    ruleFinding(
      pack,
      typefaceCheck.id,
      typeface.status,
      typeface.status === "pass"
        ? "Typeface evidence satisfies the Civil Local Rule check"
        : "Typeface evidence does not prove the Civil Local Rule check",
      typeface.evidence,
    ),
    ruleFinding(
      pack,
      spacingCheck.id,
      spacing.status,
      spacing.status === "pass"
        ? "Ordinary text is double spaced"
        : "Ordinary text spacing does not satisfy the Civil Local Rule check",
      spacing.evidence,
    ),
    ruleFinding(
      pack,
      metadataCheck.id,
      metadataPasses ? "pass" : "fail",
      metadataPasses
        ? "Required first-page matter facts are present"
        : "Required first-page matter facts are missing",
      {
        missing: missingMetadata,
        counselCount: document.metadata.counsel.length,
        incompleteCounsel,
      },
    ),
    ruleFinding(
      pack,
      footerCheck.id,
      footerPasses ? "pass" : "fail",
      footerPasses
        ? "Footer template includes case title and case number"
        : "Footer template must include {{caseName}} and {{docketNumber}}",
      { footer },
    ),
  ];
};

const applyDataDrivenChecks = (
  document: LegalDocument,
  input: ValidationInput,
  pack: UserRulePack,
): ValidationFinding[] => {
  const measurement = input.measurement;
  const profile = measurement?.deterministic.profile;
  const counted = countedEntries(document);
  const words = wordsIn(counted);
  const compliance = document.metadata.certificates.some(
    (certificate) => certificate.kind === "compliance",
  );
  return pack.checks.map((check) => {
    switch (check.kind) {
      case "length-alternative": {
        const params = check.params as Extract<
          RuleCheckParams,
          { byFilingKind: unknown }
        >;
        const filingKind = input.filingKind ?? "default";
        const limits =
          params.byFilingKind[filingKind as keyof typeof params.byFilingKind] ??
          params.byFilingKind.default;
        if (!limits)
          return ruleFinding(
            pack,
            check.id,
            "unknown",
            "No length alternative is configured for this filing kind",
            { filingKind, alternatives: [] },
          );
        const alternativeStatuses = [
          ...(limits.pages !== undefined
            ? [
                {
                  name: "pages",
                  status:
                    measurement?.pageCount === undefined
                      ? ("unknown" as const)
                      : measurement.pageCount <= limits.pages
                        ? ("pass" as const)
                        : ("fail" as const),
                },
              ]
            : []),
          ...(limits.words !== undefined
            ? [
                {
                  name: "words",
                  status:
                    limits.complianceCertificateRequired && !compliance
                      ? ("fail" as const)
                      : words <= limits.words
                        ? ("pass" as const)
                        : ("fail" as const),
                },
              ]
            : []),
          ...(limits.monospacedLines !== undefined
            ? [
                {
                  name: "monospaced-lines",
                  status:
                    measurement === undefined
                      ? ("unknown" as const)
                      : limits.complianceCertificateRequired && !compliance
                        ? ("fail" as const)
                        : measurement.deterministic.totalVisualLines <=
                            limits.monospacedLines
                          ? ("pass" as const)
                          : ("fail" as const),
                },
              ]
            : []),
        ] as const;
        const status = alternativeStatuses.some(
          (alternative) => alternative.status === "pass",
        )
          ? "pass"
          : alternativeStatuses.length > 0 &&
              alternativeStatuses.every(
                (alternative) => alternative.status === "fail",
              )
            ? "fail"
            : "unknown";
        return ruleFinding(
          pack,
          check.id,
          status,
          status === "pass"
            ? "A configured length alternative is satisfied"
            : "No configured length alternative is proved",
          {
            filingKind,
            pages: measurement?.pageCount ?? null,
            words,
            monospacedLines:
              measurement?.deterministic.totalVisualLines ?? null,
            limits,
            alternatives: alternativeStatuses,
            complianceCertificatePresent: compliance,
            countedBlockIds: counted.map((entry) => entry.id),
          },
          counted.map((entry) => entry.position),
        );
      }
      case "page-size": {
        const params = check.params as {
          widthTwips: number;
          heightTwips: number;
        };
        const status =
          profile === undefined
            ? "unknown"
            : profile.page.widthTwips === params.widthTwips &&
                profile.page.heightTwips === params.heightTwips
              ? "pass"
              : "fail";
        return ruleFinding(
          pack,
          check.id,
          status,
          status === "pass"
            ? "Page size satisfies the configured requirement"
            : "Page size does not satisfy the configured requirement",
          {
            widthTwips: profile?.page.widthTwips ?? null,
            heightTwips: profile?.page.heightTwips ?? null,
            requiredWidthTwips: params.widthTwips,
            requiredHeightTwips: params.heightTwips,
          },
        );
      }
      case "margin-minimum": {
        const params = check.params as { minimumTwips: number };
        const margins = profile?.page.marginsTwips;
        const status =
          margins === undefined
            ? "unknown"
            : Object.values(margins).every(
                  (margin) => margin >= params.minimumTwips,
                )
              ? "pass"
              : "fail";
        return ruleFinding(
          pack,
          check.id,
          status,
          status === "pass"
            ? "All page margins satisfy the configured minimum"
            : "One or more page margins are below the configured minimum",
          { marginsTwips: margins ?? null, minimumTwips: params.minimumTwips },
        );
      }
      case "typeface": {
        const params = check.params as {
          minimumTwips: number;
          mode: "proportional" | "monospaced";
          requireVerifiedPitch?: boolean;
        };
        const result = typefaceStatus(
          measurement,
          params.minimumTwips,
          params.mode,
          params.requireVerifiedPitch ?? false,
        );
        return ruleFinding(
          pack,
          check.id,
          result.status,
          result.status === "pass"
            ? "Typeface evidence satisfies the configured requirement"
            : "Typeface evidence does not prove the configured requirement",
          result.evidence,
        );
      }
      case "line-spacing": {
        const params = check.params as { doubleSpacedOrdinary?: boolean };
        const required = params.doubleSpacedOrdinary ?? true;
        const result = required
          ? spacingStatus(measurement)
          : {
              status: "pass" as const,
              evidence: { ordinaryDoubleSpaced: null, requirement: false },
            };
        return ruleFinding(
          pack,
          check.id,
          result.status,
          result.status === "pass"
            ? "Ordinary text satisfies the configured line-spacing requirement"
            : "Ordinary text does not satisfy the configured line-spacing requirement",
          result.evidence,
        );
      }
      case "counted-lines-maximum": {
        const params = check.params as { perPageMaximum: number };
        const countedLines = measurement?.deterministic.countedLinesByPage;
        const visualLines = measurement?.deterministic.visualLinesByPage;
        const lines = countedLines ?? visualLines;
        const status =
          lines === undefined
            ? "unknown"
            : lines.every((lineCount) => lineCount <= params.perPageMaximum)
              ? "pass"
              : "fail";
        return ruleFinding(
          pack,
          check.id,
          status,
          status === "pass"
            ? "Every page satisfies the configured counted-lines maximum"
            : "One or more pages exceed or cannot prove the configured counted-lines maximum",
          {
            lines: lines ?? null,
            countedLineSource: countedLines
              ? "deterministic-counted-lines"
              : "deterministic-visual-lines",
            maximum: params.perPageMaximum,
          },
        );
      }
      case "required-metadata": {
        const params = check.params as {
          fields: readonly string[];
          requireCounselComplete?: boolean;
          requireComplianceCertificate?: boolean;
        };
        const missing = params.fields.filter((field) => {
          const value =
            document.metadata[field as keyof typeof document.metadata];
          return typeof value !== "string" || value.trim().length === 0;
        });
        const incompleteCounsel =
          (params.requireCounselComplete ?? false)
            ? document.metadata.counsel
                .filter(
                  (counsel) =>
                    !counsel.name.trim() ||
                    !counsel.barNumber?.trim() ||
                    !counsel.addressLines?.some((line) => line.trim()) ||
                    !counsel.phone?.trim() ||
                    !counsel.email?.trim(),
                )
                .map((counsel) => counsel.id)
            : [];
        const complianceMissing =
          params.requireComplianceCertificate === true && !compliance;
        const status =
          missing.length === 0 &&
          (params.requireCounselComplete !== true ||
            document.metadata.counsel.length > 0) &&
          incompleteCounsel.length === 0 &&
          !complianceMissing
            ? "pass"
            : "fail";
        return ruleFinding(
          pack,
          check.id,
          status,
          status === "pass"
            ? "Required metadata is present"
            : "Required metadata is missing or incomplete",
          {
            missing,
            incompleteCounsel,
            complianceCertificatePresent: compliance,
            complianceCertificateMissing: complianceMissing,
          },
        );
      }
      case "required-block": {
        const params = check.params as { kinds: readonly string[] };
        const present = new Set<string>(
          flattenBlocks(document.blocks).map((block) => block.kind),
        );
        const missing = params.kinds.filter((kind) => !present.has(kind));
        return ruleFinding(
          pack,
          check.id,
          missing.length === 0 ? "pass" : "fail",
          missing.length === 0
            ? "Required legal blocks are present"
            : `Required legal blocks are missing: ${missing.join(", ")}`,
          { requiredKinds: params.kinds, missingKinds: missing },
        );
      }
      case "required-footer": {
        const params = check.params as { requiredTokens: readonly string[] };
        const footer = document.chrome.footers?.default ?? "";
        const missing = params.requiredTokens.filter(
          (token) => !footer.includes(token),
        );
        return ruleFinding(
          pack,
          check.id,
          missing.length === 0 ? "pass" : "fail",
          missing.length === 0
            ? "Footer contains all required tokens"
            : `Footer is missing required tokens: ${missing.join(", ")}`,
          {
            footer,
            requiredTokens: params.requiredTokens,
            missingTokens: missing,
          },
        );
      }
      case "reference-integrity": {
        const reference = referenceFinding(document);
        return ruleFinding(
          pack,
          check.id,
          reference.status,
          reference.message,
          reference.evidence,
        );
      }
    }
  });
};
export const validateLegalDocument = (
  document: LegalDocument,
  input: ValidationInput = {},
): ValidationResult => {
  const pack = input.rulePack ? builtInRulePacks[input.rulePack] : undefined;
  if (input.rulePack && !pack)
    throw new AgentDocxError(
      "RULE_PACK_INVALID",
      `Unknown rule pack: ${input.rulePack}`,
    );
  const customPacks = (input.customPacks ?? []).map((candidate, index) =>
    validateUserRulePack(candidate, `customPacks[${index}]`),
  );
  const findings = [
    metadataFinding(document),
    structureFinding(document),
    referenceFinding(document),
    ...(pack?.id === "frap-32@2024-12-01"
      ? applyFrapChecks(document, input, pack)
      : []),
    ...(pack?.id === "cand-civil@2026-05-01"
      ? applyCandChecks(document, input, pack)
      : []),
    ...customPacks.flatMap((customPack) =>
      applyDataDrivenChecks(document, input, customPack),
    ),
  ];
  const summary = findings.reduce(
    (result, finding) => ({
      ...result,
      [finding.status]: result[finding.status] + 1,
    }),
    { pass: 0, fail: 0, unknown: 0 },
  );
  const hasErrorFailure = findings.some(
    (finding) => finding.severity === "error" && finding.status === "fail",
  );
  const hasErrorUnknown = findings.some(
    (finding) => finding.severity === "error" && finding.status === "unknown",
  );
  return {
    schemaVersion: 1,
    documentId: document.documentId,
    revision: input.revision ?? null,
    rulePack: pack?.id ?? null,
    scope: {
      certification: false,
      checkedRuleIds: findings
        .filter((finding) => finding.source === "rule")
        .map((finding) => finding.checkId),
      sourceSnapshots: [
        ...(pack
          ? [
              {
                id: pack.id,
                sourceUrl: pack.sourceUrl,
                effectiveDate: pack.effectiveDate,
                sha256: pack.sourceSha256,
              },
            ]
          : []),
        ...customPacks.map((customPack) => ({
          id: customPack.id,
          sourceUrl: customPack.sourceUrl,
          effectiveDate: customPack.effectiveDate,
          sha256: customPack.sourceSha256,
        })),
      ],
      unmodeledProvisions: [
        ...(pack?.unmodeledProvisions ?? []),
        ...customPacks.flatMap((customPack) => customPack.unmodeledProvisions),
      ],
    },
    status: hasErrorFailure ? "fail" : hasErrorUnknown ? "unknown" : "pass",
    summary,
    findings,
  };
};
