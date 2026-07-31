import { indexSections, normalizeMarkdown } from "./markdown.js";
import { paginate } from "./layout/paginate.js";
import { loadFonts, resolveProfile, resolvedProfile } from "./resolve.js";
import {
  AgentDocxError,
  type Budget,
  type DeterministicResult,
  type EstimateOptions,
  type TrimOpportunity,
} from "./types.js";
export async function estimateMarkdown(
  markdown: string,
  options: EstimateOptions = {},
): Promise<DeterministicResult> {
  if (typeof markdown !== "string")
    throw new AgentDocxError("INVALID_ARGUMENT", "markdown must be a string");
  if (
    options.pageLimit !== undefined &&
    (!Number.isInteger(options.pageLimit) || options.pageLimit <= 0)
  )
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      "pageLimit must be a positive integer",
    );
  if (
    options.trim &&
    options.trim.maxCandidates !== undefined &&
    (!Number.isInteger(options.trim.maxCandidates) ||
      options.trim.maxCandidates < 1 ||
      options.trim.maxCandidates > 100)
  )
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      "trim.maxCandidates must be an integer from 1 through 100",
    );
  if (
    options.trim &&
    options.trim.maxLastLineRatio !== undefined &&
    (!Number.isFinite(options.trim.maxLastLineRatio) ||
      options.trim.maxLastLineRatio < 0 ||
      options.trim.maxLastLineRatio > 1)
  )
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      "trim.maxLastLineRatio must be from 0 through 1",
    );
  const profile = resolveProfile(options);
  const selectedLimit =
    options.pageLimit ??
    (options.filingKind
      ? profile.filingPageLimits[options.filingKind]
      : undefined);
  const fonts = await loadFonts(options.fontSet, profile.requestedFontFamily);
  const document = normalizeMarkdown(markdown);
  const sectionIndex =
    options.sectionDiagnostics === true
      ? indexSections(document.blocks)
      : undefined;
  const layout = paginate(document, profile, fonts, sectionIndex);
  const warnings = [
    ...profile.warnings,
    ...(options.template?.warnings ?? []),
    ...fonts.warnings,
    ...layout.warnings,
  ];
  const result: DeterministicResult = {
    schemaVersion: 1,
    pageCount: layout.pageCount,
    equivalentPages: layout.equivalentPages,
    totalVisualLines: layout.totalVisualLines,
    visualLinesByPage: layout.visualLinesByPage,
    lastPage: layout.lastPage
      ? { source: "deterministic", ...layout.lastPage }
      : null,
    profile: resolvedProfile(profile, fonts, options),
    warnings,
  };
  if (layout.sections) result.sections = layout.sections;
  if (options.paragraphDiagnostics || options.trim)
    result.paragraphs = layout.paragraphs;
  if (options.trim) {
    const threshold = options.trim.maxLastLineRatio ?? 0.35;
    const max = options.trim.maxCandidates ?? 10;
    const candidates = layout.paragraphs
      .filter(
        (
          paragraph,
        ): paragraph is typeof paragraph & {
          oneLineReduction: NonNullable<(typeof paragraph)["oneLineReduction"]>;
        } =>
          paragraph.oneLineReduction !== null &&
          paragraph.lastLineRatio <= threshold,
      )
      .sort(
        (a, b) =>
          a.oneLineReduction.estimatedRemovalTwips -
            b.oneLineReduction.estimatedRemovalTwips ||
          a.position.start.offset - b.position.start.offset,
      )
      .slice(0, max);
    result.trimOpportunities = candidates.map(
      (paragraph, index): TrimOpportunity => ({
        ...paragraph,
        rank: index + 1,
        message:
          "This block may lose one wrapped line after removing or rephrasing approximately the reported width; verify by re-running pagination.",
      }),
    );
  }
  if (selectedLimit !== undefined) {
    const pitch =
      result.lastPage && result.lastPage.bodyLineCapacity > 0
        ? result.lastPage.usableTwips / result.lastPage.bodyLineCapacity
        : profile.body.fontSizeTwips;
    const bodyRemaining =
      ((selectedLimit - result.equivalentPages) *
        (result.lastPage?.usableTwips ??
          profile.page.heightTwips -
            profile.page.marginsTwips.top -
            profile.page.marginsTwips.bottom)) /
      pitch;
    const budget: Budget = {
      limitPages: selectedLimit,
      withinLimit: result.pageCount <= selectedLimit,
      pagesRemaining: selectedLimit - result.pageCount,
      equivalentPagesRemaining: selectedLimit - result.equivalentPages,
      bodyLineEquivalentsRemaining: bodyRemaining,
      fractionalFieldsSource: "deterministic",
    };
    result.budget = budget;
    if (result.sections) {
      result.sections = result.sections.map((section) => {
        const pagesBeyondLimit = section.pages
          .map((page) => page.page)
          .filter((page) => page > selectedLimit);
        return {
          ...section,
          pageBudget: {
            limitPages: selectedLimit,
            withinLimit: pagesBeyondLimit.length === 0,
            pagesBeyondLimit,
          },
        };
      });
    }
  }
  return result;
}
