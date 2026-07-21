import type { LayoutProfile, TextStyle, BuiltInProfileId } from "./types.js";
const double = { rule: "auto", numerator: 480, denominator: 240 } as const;
const single = { rule: "auto", numerator: 240, denominator: 240 } as const;
const style = (
  fontSizeTwips: number,
  lineSpacing: TextStyle["lineSpacing"],
  extra: Partial<TextStyle> = {},
): TextStyle => ({
  fontSizeTwips,
  bold: false,
  italic: false,
  lineSpacing,
  beforeTwips: 0,
  afterTwips: 0,
  leftIndentTwips: 0,
  rightIndentTwips: 0,
  firstLineIndentTwips: 0,
  hangingIndentTwips: 0,
  keepWithNext: false,
  keepLines: false,
  ...extra,
});
const headings = (s: TextStyle) =>
  Object.freeze({
    "1": { ...s },
    "2": { ...s },
    "3": { ...s },
    "4": { ...s },
    "5": { ...s },
    "6": { ...s },
  });
const page = {
  widthTwips: 12240,
  heightTwips: 15840,
  marginsTwips: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
  headerTwips: 720,
  footerTwips: 720,
  gutterTwips: 0,
};
const conventional: LayoutProfile = {
  id: "us-district-conventional",
  label: "U.S. district court conventional estimate",
  effectiveDate: null,
  sourceUrl: null,
  sourceCitation: "Package convention; not a court rule.",
  page,
  requestedFontFamily: "Times New Roman",
  body: style(240, double),
  headings: headings(style(240, double, { bold: true, keepWithNext: true })),
  blockquote: style(240, double, {
    leftIndentTwips: 720,
    rightIndentTwips: 720,
  }),
  list: style(240, double, { leftIndentTwips: 720, hangingIndentTwips: 360 }),
  footnote: style(240, double),
  pagination: {
    widowOrphanControl: true,
    widowLines: 2,
    orphanLines: 2,
    maxCountedLinesPerPage: null,
    lineCapExclusions: [],
  },
  maxCharactersPerInch: null,
  filingPageLimits: {},
  provenance: {
    "": { source: "package", detail: "Conventional product baseline" },
  },
  warnings: [
    {
      code: "CONVENTIONAL_PROFILE_ONLY",
      severity: "warning",
      message:
        "Conventional estimate only. Federal district courts, local rules, assigned-judge standing orders, and case-specific orders vary; this profile does not certify filing compliance.",
    },
  ],
};
const frap: LayoutProfile = {
  ...conventional,
  id: "frap-32",
  label: "Federal Rule of Appellate Procedure 32",
  effectiveDate: "2024-12-01",
  sourceUrl:
    "https://www.uscourts.gov/forms-rules/current-rules-practice-procedure/federal-rules-appellate-procedure",
  sourceCitation: "Fed. R. App. P. 32(a)(4)-(7)",
  body: style(280, double),
  headings: headings(style(280, single, { bold: true, keepWithNext: true })),
  blockquote: style(280, double, {
    leftIndentTwips: 720,
    rightIndentTwips: 720,
  }),
  list: style(280, double, { leftIndentTwips: 720, hangingIndentTwips: 360 }),
  footnote: style(280, single),
  filingPageLimits: { "principal-brief": 30, "reply-brief": 15 },
  provenance: { "": { source: "rule", citation: "Fed. R. App. P. 32" } },
  warnings: [],
};
const cand: LayoutProfile = {
  ...conventional,
  id: "cand-civil",
  label: "N.D. California Civil Local Rules",
  effectiveDate: "2025-12-01",
  sourceUrl: "https://www.cand.uscourts.gov/rules/civil-local-rules/",
  sourceCitation: "N.D. Cal. Civ. L.R. 3-4(c)(2), 7-2(b), 7-3(a), 7-3(c)",
  headings: headings(style(240, double, { bold: true, keepWithNext: true })),
  blockquote: style(240, single, {
    leftIndentTwips: 720,
    rightIndentTwips: 720,
  }),
  footnote: style(240, single),
  pagination: {
    widowOrphanControl: true,
    widowLines: 2,
    orphanLines: 2,
    maxCountedLinesPerPage: 28,
    lineCapExclusions: ["footnote", "blockquote"],
  },
  maxCharactersPerInch: 10,
  filingPageLimits: {
    "motion-document": 25,
    "opposition-text": 25,
    "reply-text": 15,
  },
  provenance: {
    "": { source: "rule", citation: "N.D. Cal. Civ. L.R. 3-4(c)(2)" },
    "/page/marginsTwips": {
      source: "package",
      detail: "Rule does not specify margins; package convention",
    },
  },
  warnings: [
    {
      code: "CAND_CPI_NOT_AUTOMATICALLY_VALIDATED",
      severity: "warning",
      message:
        "The rule's proportional-font characters-per-inch predicate is not automatically validated; verify compliance independently.",
    },
  ],
};
const freeze = <T>(v: T): T => {
  if (v && typeof v === "object") {
    Object.freeze(v);
    for (const x of Object.values(v as Record<string, unknown>)) freeze(x);
  }
  return v;
};
export const builtInProfiles: Readonly<
  Record<BuiltInProfileId, LayoutProfile>
> = freeze({
  "us-district-conventional": conventional,
  "frap-32": frap,
  "cand-civil": cand,
});
