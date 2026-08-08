import type { DocumentChrome } from "../legal/model.js";
import { AgentDocxError } from "../types.js";
import type { LayoutProfile, TextStyle } from "./profile.js";

export type PageChromeKind = "default" | "first" | "even";

export type BodyBounds = {
  kind: PageChromeKind;
  headerHeightTwips: number;
  footerHeightTwips: number;
  bodyTopTwips: number;
  bodyBottomTwips: number;
  usableHeightTwips: number;
};

const textForMeasurement = (template: string, fieldDigits: number): string =>
  template.replace(
    /\{\{(?:caseName|docketNumber|documentTitle|page|pages)\}\}/g,
    (token) =>
      token === "{{page}}" || token === "{{pages}}"
        ? "9".repeat(fieldDigits)
        : "88",
  );

const linePitch = (style: TextStyle): number => {
  if (style.lineSpacing.rule === "auto")
    return Math.max(
      style.fontSizeTwips,
      Math.round((style.fontSizeTwips * style.lineSpacing.numerator) / 240),
    );
  return Math.max(style.fontSizeTwips, style.lineSpacing.twips);
};

const renderedHeight = (
  template: string | undefined,
  profile: LayoutProfile,
  fieldDigits: number,
): number => {
  if (!template) return 0;
  const text = textForMeasurement(template, fieldDigits);
  const width =
    profile.page.widthTwips -
    profile.page.marginsTwips.left -
    profile.page.marginsTwips.right -
    profile.page.gutterTwips;
  const averageGlyphWidth = Math.max(
    1,
    Math.round(profile.body.fontSizeTwips * 0.5),
  );
  const lineCapacity = Math.max(1, Math.floor(width / averageGlyphWidth));
  const lines = text
    .split(/\r?\n/)
    .reduce(
      (total, line) =>
        total + Math.max(1, Math.ceil(line.length / lineCapacity)),
      0,
    );
  return lines * linePitch(profile.body);
};

const templateFor = (
  templates: DocumentChrome["headers"] | DocumentChrome["footers"] | undefined,
  kind: PageChromeKind,
): string | undefined => templates?.[kind] ?? templates?.default;

/**
 * Returns Word-equivalent body bounds for one header/footer page kind. The same
 * calculation is shared by deterministic pagination and DOCX section margins.
 */
const bodyBoundsFor = (
  profile: LayoutProfile,
  chrome: DocumentChrome | undefined,
  kind: PageChromeKind,
  fieldDigits = 2,
): BodyBounds => {
  const headerTemplate = templateFor(chrome?.headers, kind);
  const footerTemplate = templateFor(chrome?.footers, kind);
  const headerNumber = chrome?.pageNumber?.story === "header";
  const footerNumber = chrome?.pageNumber?.story === "footer";
  const headerHeightTwips =
    renderedHeight(headerTemplate, profile, fieldDigits) +
    (headerNumber ? linePitch(profile.body) : 0);
  const footerHeightTwips =
    renderedHeight(footerTemplate, profile, fieldDigits) +
    (footerNumber ? linePitch(profile.body) : 0);
  const bodyTopTwips =
    headerTemplate || headerNumber
      ? Math.max(
          profile.page.marginsTwips.top,
          profile.page.headerTwips + headerHeightTwips,
        )
      : profile.page.marginsTwips.top;
  const bodyBottomTwips =
    footerTemplate || footerNumber
      ? Math.min(
          profile.page.heightTwips - profile.page.marginsTwips.bottom,
          profile.page.heightTwips -
            profile.page.footerTwips -
            footerHeightTwips,
        )
      : profile.page.heightTwips - profile.page.marginsTwips.bottom;
  const usableHeightTwips = bodyBottomTwips - bodyTopTwips;
  if (usableHeightTwips <= 0)
    throw new AgentDocxError(
      "INVALID_LAYOUT",
      `Header/footer chrome leaves no usable body height for ${kind} pages`,
      {
        bodyTopTwips,
        bodyBottomTwips,
        headerHeightTwips,
        footerHeightTwips,
      },
    );
  return {
    kind,
    headerHeightTwips,
    footerHeightTwips,
    bodyTopTwips,
    bodyBottomTwips,
    usableHeightTwips,
  };
};

/** Uses the tightest page-kind body box until page-kind pagination is requested. */
export const conservativeBodyBounds = (
  profile: LayoutProfile,
  chrome: DocumentChrome | undefined,
  fieldDigits = 2,
): BodyBounds =>
  ["default", "first", "even"]
    .map((kind) =>
      bodyBoundsFor(profile, chrome, kind as PageChromeKind, fieldDigits),
    )
    .sort(
      (left, right) =>
        left.usableHeightTwips - right.usableHeightTwips ||
        (left.kind < right.kind ? -1 : left.kind > right.kind ? 1 : 0),
    )[0]!;
