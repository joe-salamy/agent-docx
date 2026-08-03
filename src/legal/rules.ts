import {
  AgentDocxError,
  type FilingKind,
  type JsonValue,
  type MeasurementResult,
  type SourcePosition,
} from "../types.js";
import type {
  LegalBlock,
  LegalDocument,
  RevisionId,
  RulePackId,
} from "./model.js";

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
      id: RulePackId;
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

export type BuiltInRulePack = {
  id: RulePackId;
  sourceUrl: string;
  effectiveDate: string;
  sourceSha256: `sha256:${string}`;
  sourceExcerpt: string;
  checks: readonly { id: string; kind: RuleCheckKind; citation: string; predicate: string }[];
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
      { id: "frap32.length.principal", kind: "length-alternative", citation: "Fed. R. App. P. 32(a)(7)", predicate: "principal <= 30 pages OR <= 13000 words OR <= 1300 monospaced lines; non-page alternatives require a Rule 32(g) compliance certificate" },
      { id: "frap32.length.reply", kind: "length-alternative", citation: "Fed. R. App. P. 32(a)(7)", predicate: "reply <= 15 pages OR <= 6500 words OR <= 650 monospaced lines; non-page alternatives require a Rule 32(g) compliance certificate" },
      { id: "frap32.page-size", kind: "page-size", citation: "Fed. R. App. P. 32(a)(4)", predicate: "page is exactly 8.5 by 11 inches" },
      { id: "frap32.margin", kind: "margin-minimum", citation: "Fed. R. App. P. 32(a)(4)", predicate: "each margin is at least one inch" },
      { id: "frap32.typeface.proportional", kind: "typeface", citation: "Fed. R. App. P. 32(a)(5)", predicate: "proportional serif text is at least 14 point" },
      { id: "frap32.typeface.monospaced", kind: "typeface", citation: "Fed. R. App. P. 32(a)(5)", predicate: "monospaced text contains no more than 10.5 characters per inch" },
      { id: "frap32.spacing", kind: "line-spacing", citation: "Fed. R. App. P. 32(a)(4)", predicate: "ordinary text is double spaced" },
    ],
    unmodeledProvisions: [
      "Rule 32(a)(1)-(3), (6), and formatting conditions not represented by the legal IR",
      "Court-specific local-rule exclusions beyond an explicitly cited directive",
    ],
  },
  "cand-civil@2026-05-01": {
    id: "cand-civil@2026-05-01",
    sourceUrl: "https://cand.uscourts.gov/rules-forms-fees/local-rules/civil-local-rules",
    effectiveDate: "2026-05-01",
    sourceSha256:
      "sha256:82383622a53a6ffe15bdc7931c263db4164c2e4a263a79fe4e53eb80d86dcbb5",
    sourceExcerpt: candSource,
    checks: [
      { id: "cand.length.motion", kind: "length-alternative", citation: "Civil L.R. 7-2(b)", predicate: "motion document <= 25 pages" },
      { id: "cand.length.opposition", kind: "length-alternative", citation: "Civil L.R. 7-3(a), 7-3(c)", predicate: "opposition text <= 25 pages" },
      { id: "cand.length.reply", kind: "length-alternative", citation: "Civil L.R. 7-4(b)", predicate: "reply text <= 15 pages" },
      { id: "cand.lines", kind: "counted-lines-maximum", citation: "Civil L.R. 3-4(c)(2)", predicate: "each page has at most 28 counted lines" },
      { id: "cand.typeface", kind: "typeface", citation: "Civil L.R. 3-4(c)(2)", predicate: "all text is at least 12 point in a verified proportional serif face" },
      { id: "cand.spacing", kind: "line-spacing", citation: "Civil L.R. 3-4(c)(2)", predicate: "ordinary text is double spaced" },
      { id: "cand.first-page", kind: "required-metadata", citation: "Civil L.R. 3-4(a)", predicate: "counsel, court, case, and document-title facts are present" },
      { id: "cand.footer", kind: "required-footer", citation: "Civil L.R. 3-4(c)(3)", predicate: "footer contains case title and case number" },
    ],
    unmodeledProvisions: [
      "Judge-specific and conditional legends",
      "Individual judge standing orders and page-limit exceptions",
    ],
  },
};

export type ValidationInput = {
  revision?: RevisionId | null;
  rulePack?: RulePackId;
  filingKind?: FilingKind;
  measurement?: Omit<MeasurementResult, "generatedDocx">;
};
type CountedEntry = {
  id: LegalBlock["id"];
  kind: string;
  text: string;
  position: SourcePosition;
};

const textFor = (block: LegalBlock): string => {
  if ("runs" in block) return block.runs.map((run) => run.text).join("");
  if (block.kind === "table")
    return block.rows
      .flatMap((row) => row.flatMap((cell) => cell.paragraphs))
      .flatMap((paragraph) => paragraph.runs)
      .map((run) => run.text)
      .join(" ");
  if (block.kind === "list")
    return block.items
      .flatMap((item) => item.paragraphs)
      .flatMap((paragraph) => paragraph.runs)
      .map((run) => run.text)
      .join(" ");
  return block.sourceText;
};

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
  const visit = (blocks: readonly LegalBlock[], inheritedExclusion = false): void => {
    for (const block of blocks) {
      const excludedHere = inheritedExclusion || block.kind === "length-exclusion";
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
  spacing: { rule: "auto"; numerator: number; denominator: 240 } | { rule: "exact" | "atLeast"; twips: number },
): boolean =>
  spacing.rule === "auto"
    ? spacing.numerator >= 480
    : spacing.twips >= 480;

const ruleFinding = (
  pack: BuiltInRulePack,
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
  citation: pack.checks.find((check) => check.id === checkId)?.citation,
  sourceUrl: pack.sourceUrl,
  effectiveDate: pack.effectiveDate,
  ...(positions && positions.length > 0 ? { positions } : {}),
  evidence: {
    sourceSnapshotId: pack.id,
    sourceSha256: pack.sourceSha256,
    predicate: pack.checks.find((check) => check.id === checkId)?.predicate ?? null,
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
    missing.length === 0 ? "Core matter metadata is present" : `Missing core metadata: ${missing.join(", ")}`,
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
  const duplicateCounsel = duplicates(document.metadata.counsel.map((entry) => entry.id));
  const duplicateCertificates = duplicates(
    document.metadata.certificates.map((entry) => entry.id),
  );
  const duplicateAnnotations = duplicates(
    document.annotations.map((annotation) => annotation.id),
  );
  const counsel = new Set(document.metadata.counsel.map((entry) => entry.id));
  const certificates = new Set(document.metadata.certificates.map((entry) => entry.id));
  const footnotes = new Set(document.footnotes.map((footnote) => footnote.label));
  const unresolvedBlocks = blocks.flatMap((block) => {
    if (block.kind === "signature" && !counsel.has(block.counselId)) return [block.id];
    if (block.kind === "certificate" && !certificates.has(block.certificateId)) return [block.id];
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
  const kinds = new Set(flattenBlocks(document.blocks).map((block) => block.kind));
  const required = ["caption", "signature", "certificate"] as const;
  const missing = required.filter((kind) => !kinds.has(kind));
  return packageFinding(
    "package.structure",
    missing.length === 0 ? "pass" : "fail",
    "warning",
    missing.length === 0 ? "Expected filing structure is present" : `Product-convention blocks are absent: ${missing.join(", ")}`,
    { missing },
  );
};

const typefaceStatus = (
  measurement: ValidationInput["measurement"],
  minimumTwips: number,
  mode: "proportional" | "monospaced",
  requireVerifiedPitch = false,
): { status: ValidationFinding["status"]; evidence: Readonly<Record<string, JsonValue>> } => {
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
    classification: monospaced ? "monospaced" : serif ? "proportional-serif" : "unknown",
    sizes,
    minimumTwips,
    charactersPerInch: profile.maxCharactersPerInch,
  };
  if (mode === "monospaced") {
    if (!monospaced)
      return { status: "unknown", evidence: { ...evidence, reason: "Font is not verified monospaced" } };
    if (profile.maxCharactersPerInch === null)
      return { status: "unknown", evidence: { ...evidence, reason: "Characters-per-inch is unproved" } };
    return {
      status: profile.maxCharactersPerInch <= 10.5 ? "pass" : "fail",
      evidence,
    };
  }
  if (monospaced || !serif)
    return {
      status: "unknown",
      evidence: { ...evidence, reason: "Proportional serif classification is unproved" },
    };
  if (!sizePasses) return { status: "fail", evidence };
  if (requireVerifiedPitch)
    return {
      status: "unknown",
      evidence: {
        ...evidence,
        pitchVerified: false,
        manualVerificationRequired: true,
        reason: "The selected font's applicable pitch is not independently verified",
      },
    };
  return { status: "pass", evidence };
};

const spacingStatus = (
  measurement: ValidationInput["measurement"],
): { status: ValidationFinding["status"]; evidence: Readonly<Record<string, JsonValue>> } => {
  if (!measurement)
    return { status: "unknown", evidence: { reason: "No deterministic measurement was supplied" } };
  const profile = measurement.deterministic.profile;
  const ordinary = [profile.body.lineSpacing, profile.list.lineSpacing];
  return {
    status: ordinary.every(doubleSpaced) ? "pass" : "fail",
    evidence: { ordinaryDoubleSpaced: ordinary.map(doubleSpaced) },
  };
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
  const reply = input.filingKind === "reply-brief";
  const lengthCheck = reply ? "frap32.length.reply" : "frap32.length.principal";
  const limits = reply
    ? { pages: 15, words: 6500, lines: 650 }
    : { pages: 30, words: 13000, lines: 1300 };
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
    { name: "pages", status: pages === null ? "unknown" : pages <= limits.pages ? "pass" : "fail" },
    {
      name: "words",
      status:
        complianceBlock === undefined
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
          : complianceBlock === undefined
            ? "fail"
            : lines <= limits.lines
              ? "pass"
              : "fail",
    },
  ] as const;
  const lengthStatus = alternatives.some((alternative) => alternative.status === "pass")
    ? "pass"
    : alternatives.every((alternative) => alternative.status === "fail")
      ? "fail"
      : "unknown";
  const sizePasses =
    profile?.page.widthTwips === 12240 && profile.page.heightTwips === 15840;
  const margins = profile?.page.marginsTwips;
  const marginPasses = margins
    ? Object.values(margins).every((margin) => margin >= 1440)
    : false;
  const monospaced = /mono|courier|consolas|menlo/i.test(
    profile?.requestedFontFamily ?? "",
  );
  const typeface = typefaceStatus(
    measurement,
    280,
    monospaced ? "monospaced" : "proportional",
  );
  const typefaceCheck = monospaced
    ? "frap32.typeface.monospaced"
    : "frap32.typeface.proportional";
  const spacing = spacingStatus(measurement);
  return [
    ruleFinding(
      pack,
      lengthCheck,
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
    ruleFinding(pack, "frap32.page-size", profile ? (sizePasses ? "pass" : "fail") : "unknown", sizePasses ? "Page size is 8.5 by 11 inches" : "Page size is not verified as 8.5 by 11 inches", { widthTwips: profile?.page.widthTwips ?? null, heightTwips: profile?.page.heightTwips ?? null }),
    ruleFinding(pack, "frap32.margin", profile ? (marginPasses ? "pass" : "fail") : "unknown", marginPasses ? "All page margins are at least one inch" : "One or more page margins are below one inch", { marginsTwips: margins ?? null, minimumTwips: 1440 }),
    ruleFinding(pack, typefaceCheck, typeface.status, typeface.status === "pass" ? "Typeface evidence satisfies the selected Rule 32 check" : "Typeface evidence does not prove the selected Rule 32 check", typeface.evidence),
    ruleFinding(pack, "frap32.spacing", spacing.status, spacing.status === "pass" ? "Ordinary text is double spaced" : "Ordinary text spacing does not satisfy Rule 32", spacing.evidence),
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
  const length = kind === "motion-document"
    ? { id: "cand.length.motion", limit: 25 }
    : kind === "opposition-text"
      ? { id: "cand.length.opposition", limit: 25 }
      : { id: "cand.length.reply", limit: 15 };
  const typeface = typefaceStatus(measurement, 240, "proportional", true);
  const spacing = spacingStatus(measurement);
  const lines = measurement?.deterministic.visualLinesByPage ?? null;
  const linesStatus = lines === null
    ? "unknown"
    : lines.every((lineCount) => lineCount <= 28)
      ? "pass"
      : "fail";
  const coreMetadata = [
    ["court", document.metadata.court],
    ["jurisdiction", document.metadata.jurisdiction],
    ["caseName", document.metadata.caseName],
    ["docketNumber", document.metadata.docketNumber],
    ["documentTitle", document.metadata.documentTitle],
  ] as const;
  const missingMetadata = coreMetadata
    .filter(([, value]) => value.trim().length === 0)
    .map(([key]) => key);
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
  const footer = document.chrome.footers?.default ?? "";
  const footerPasses =
    footer.includes("{{caseName}}") && footer.includes("{{docketNumber}}");
  return [
    ruleFinding(pack, length.id, pages === null ? "unknown" : pages <= length.limit ? "pass" : "fail", pages !== null && pages <= length.limit ? "Document page count is within the selected Civil Local Rule limit" : "Document page count exceeds or cannot prove the selected Civil Local Rule limit", { filingKind: kind ?? "reply-text", totalPages: pages, pagesOfBriefText: pages, limit: length.limit }),
    ruleFinding(pack, "cand.lines", linesStatus, linesStatus === "pass" ? "Each page has at most 28 counted lines" : "One or more pages exceed or cannot prove the 28-line limit", { visualLinesByPage: lines ?? [], countedLineSource: "deterministic-visual-lines", maximum: 28 }),
    ruleFinding(pack, "cand.typeface", typeface.status, typeface.status === "pass" ? "Typeface evidence satisfies the Civil Local Rule check" : "Typeface evidence does not prove the Civil Local Rule check", typeface.evidence),
    ruleFinding(pack, "cand.spacing", spacing.status, spacing.status === "pass" ? "Ordinary text is double spaced" : "Ordinary text spacing does not satisfy the Civil Local Rule check", spacing.evidence),
    ruleFinding(
      pack,
      "cand.first-page",
      missingMetadata.length === 0 &&
        document.metadata.counsel.length > 0 &&
        incompleteCounsel.length === 0
        ? "pass"
        : "fail",
      missingMetadata.length === 0 &&
        document.metadata.counsel.length > 0 &&
        incompleteCounsel.length === 0
        ? "Required first-page matter facts are present"
        : "Required first-page matter facts are missing",
      {
        missingMetadata,
        counselCount: document.metadata.counsel.length,
        incompleteCounsel,
      },
    ),
    ruleFinding(pack, "cand.footer", footerPasses ? "pass" : "fail", footerPasses ? "Footer template includes case title and case number" : "Footer template must include {{caseName}} and {{docketNumber}}", { footer }),
  ];
};

export const validateLegalDocument = (
  document: LegalDocument,
  input: ValidationInput = {},
): ValidationResult => {
  const pack = input.rulePack ? builtInRulePacks[input.rulePack] : undefined;
  if (input.rulePack && !pack)
    throw new AgentDocxError("RULE_PACK_INVALID", `Unknown rule pack: ${input.rulePack}`);
  const findings = [
    metadataFinding(document),
    structureFinding(document),
    referenceFinding(document),
    ...(pack?.id === "frap-32@2024-12-01" ? applyFrapChecks(document, input, pack) : []),
    ...(pack?.id === "cand-civil@2026-05-01" ? applyCandChecks(document, input, pack) : []),
  ];
  const summary = findings.reduce(
    (result, finding) => ({ ...result, [finding.status]: result[finding.status] + 1 }),
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
      sourceSnapshots: pack
        ? [{
            id: pack.id,
            sourceUrl: pack.sourceUrl,
            effectiveDate: pack.effectiveDate,
            sha256: pack.sourceSha256,
          }]
        : [],
      unmodeledProvisions: pack?.unmodeledProvisions ?? [],
    },
    status: hasErrorFailure ? "fail" : hasErrorUnknown ? "unknown" : "pass",
    summary,
    findings,
  };
};

export const emptyValidationResult = (
  document: LegalDocument,
  revision: RevisionId | null = null,
): ValidationResult => ({
  schemaVersion: 1,
  documentId: document.documentId,
  revision,
  rulePack: null,
  scope: {
    certification: false,
    checkedRuleIds: [],
    sourceSnapshots: [],
    unmodeledProvisions: [],
  },
  status: "pass",
  summary: { pass: 0, fail: 0, unknown: 0 },
  findings: [],
});
