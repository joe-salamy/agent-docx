import {
  decodeDocxXml as decodeXml,
  docxXmlAttribute as attr,
  parseDocxXml as parse,
  readDocxParts,
  resolveOpcTarget,
  sha256Hex,
} from "./package.js";
import {
  isWordprocessingElement,
  parseRelationships,
  relationshipPartFor,
  type Relationship,
} from "./opc.js";
import { codePointCompare } from "./helpers.js";
import { builtInProfiles } from "../profiles.js";
import { AgentDocxError } from "../types.js";
import type { Diagnostic } from "../types.js";
import type {
  LayoutProfile,
  PageGeometry,
  TextStyle,
} from "../layout/profile.js";
import type {
  DocxTemplateInspection,
  InspectTemplateOptions,
} from "./contracts.js";

export { DOCX_LIMITS } from "./package.js";

type Parts = Record<string, Uint8Array>;
type LoadedParts = { parts: Parts; names: readonly string[] };
type HeaderFooterReference = {
  kind: "header" | "footer";
  variant: "default" | "first" | "even";
  relationshipId: string;
};
type TemplateSection = {
  page: PageGeometry;
  headerFooterReferences: readonly HeaderFooterReference[];
};

const readParts = async (bytes: Uint8Array): Promise<LoadedParts> => {
  const names: string[] = [];
  const loaded = await readDocxParts(bytes, (name) => {
    names.push(name);
    return (
      name === "[Content_Types].xml" ||
      name.endsWith(".rels") ||
      (name.startsWith("word/") && name.endsWith(".xml"))
    );
  });
  return { parts: Object.fromEntries(loaded), names: names.sort() };
};

const relationships = (
  xml: string | undefined,
  sourcePart: string,
): readonly Relationship[] => (xml ? parseRelationships(xml, sourcePart) : []);

const sourcePartForRels = (partPath: string): string => {
  if (partPath === "_rels/.rels") return "";
  const components = partPath.split("/");
  const name = components.pop()!.slice(0, -".rels".length);
  components.pop();
  return [...components, name].join("/");
};

function sectionGeometry(
  xml: string,
  fallback: PageGeometry,
): readonly TemplateSection[] {
  const sections: TemplateSection[] = [];
  let page: PageGeometry | null = null;
  let references: HeaderFooterReference[] = [];
  parse(
    xml,
    (tag) => {
      if (!isWordprocessingElement(tag)) return;
      if (tag.local === "sectPr") {
        page = structuredClone(fallback);
        references = [];
      } else if (page && tag.local === "pgSz") {
        const width = Number(attr(tag, "w")),
          height = Number(attr(tag, "h"));
        if (Number.isFinite(width) && width > 0) page.widthTwips = width;
        if (Number.isFinite(height) && height > 0) page.heightTwips = height;
      } else if (page && tag.local === "pgMar") {
        for (const key of ["top", "right", "bottom", "left"] as const) {
          const value = Number(attr(tag, key));
          if (Number.isFinite(value)) page.marginsTwips[key] = value;
        }
        for (const [xmlKey, key] of [
          ["header", "headerTwips"],
          ["footer", "footerTwips"],
          ["gutter", "gutterTwips"],
        ] as const) {
          const value = Number(attr(tag, xmlKey));
          if (Number.isFinite(value) && value >= 0) page[key] = value;
        }
      } else if (
        page &&
        (tag.local === "headerReference" || tag.local === "footerReference")
      ) {
        const relationshipId = attr(tag, "id");
        const variant = attr(tag, "type");
        if (
          !relationshipId ||
          (variant !== "default" && variant !== "first" && variant !== "even")
        )
          throw new AgentDocxError(
            "DOCX_INVALID",
            "Section header or footer reference is invalid",
          );
        references.push({
          kind: tag.local === "headerReference" ? "header" : "footer",
          variant,
          relationshipId,
        });
      }
    },
    (tag) => {
      if (!isWordprocessingElement(tag)) return;
      if (tag.local === "sectPr" && page) {
        sections.push({ page, headerFooterReferences: references });
        page = null;
      }
    },
  );
  if (sections.length === 0)
    sections.push({
      page: structuredClone(fallback),
      headerFooterReferences: [],
    });
  return sections;
}
function fallbackStyle(style: TextStyle, family: string, id: string) {
  return {
    styleId: id,
    name: id,
    resolved: structuredClone(style),
    requestedFontFamily: family,
    provenance: { "": "fallback" },
  } as const;
}

type StylePatch = {
  id: string;
  name: string | null;
  basedOn: string | null;
  fontFamily: string | null;
  values: Partial<TextStyle>;
};

const booleanValue = (value: string | undefined): boolean =>
  value === undefined || !/^(?:0|false|off)$/i.test(value);

const themeFonts = (
  xml: string | undefined,
): Readonly<Record<"major" | "minor", string | null>> => {
  const result: Record<"major" | "minor", string | null> = {
    major: null,
    minor: null,
  };
  if (!xml) return result;
  let scope: "major" | "minor" | null = null;
  parse(
    xml,
    (tag) => {
      if (tag.local === "majorFont") scope = "major";
      else if (tag.local === "minorFont") scope = "minor";
      else if (scope && tag.local === "latin") {
        const family = attr(tag, "typeface");
        if (family) result[scope] = family;
      }
    },
    (tag) => {
      if (
        (tag.local === "majorFont" && scope === "major") ||
        (tag.local === "minorFont" && scope === "minor")
      )
        scope = null;
    },
  );
  return result;
};

const parseStyles = (
  xml: string | undefined,
  theme: Readonly<Record<"major" | "minor", string | null>>,
): ReadonlyMap<string, StylePatch> => {
  const styles = new Map<string, StylePatch>();
  if (!xml) return styles;
  let current: StylePatch | null = null;
  parse(
    xml,
    (tag) => {
      if (tag.local === "style") {
        const id = attr(tag, "styleId");
        const type = attr(tag, "type");
        if (!id || (type !== undefined && type !== "paragraph")) return;
        current = {
          id,
          name: null,
          basedOn: null,
          fontFamily: null,
          values: {},
        };
        return;
      }
      if (!current) return;
      if (tag.local === "name") current.name = attr(tag, "val") ?? null;
      else if (tag.local === "basedOn")
        current.basedOn = attr(tag, "val") ?? null;
      else if (tag.local === "rFonts") {
        const direct =
          attr(tag, "ascii") ??
          attr(tag, "hAnsi") ??
          attr(tag, "cs") ??
          attr(tag, "eastAsia");
        const themed =
          attr(tag, "asciiTheme") ??
          attr(tag, "hAnsiTheme") ??
          attr(tag, "csTheme") ??
          attr(tag, "eastAsiaTheme");
        const group =
          themed?.toLowerCase().startsWith("major") === true
            ? "major"
            : themed?.toLowerCase().startsWith("minor") === true
              ? "minor"
              : null;
        current.fontFamily = direct ?? (group ? theme[group] : null);
      } else if (tag.local === "sz") {
        const halfPoints = Number(attr(tag, "val"));
        if (Number.isFinite(halfPoints) && halfPoints > 0)
          current.values.fontSizeTwips = Math.round(halfPoints * 10);
      } else if (tag.local === "b")
        current.values.bold = booleanValue(attr(tag, "val"));
      else if (tag.local === "i")
        current.values.italic = booleanValue(attr(tag, "val"));
      else if (tag.local === "keepNext")
        current.values.keepWithNext = booleanValue(attr(tag, "val"));
      else if (tag.local === "keepLines")
        current.values.keepLines = booleanValue(attr(tag, "val"));
      else if (tag.local === "spacing") {
        const before = Number(attr(tag, "before"));
        const after = Number(attr(tag, "after"));
        const line = Number(attr(tag, "line"));
        const rule = attr(tag, "lineRule")?.toLowerCase();
        if (Number.isFinite(before) && before >= 0)
          current.values.beforeTwips = before;
        if (Number.isFinite(after) && after >= 0)
          current.values.afterTwips = after;
        if (Number.isFinite(line) && line > 0)
          current.values.lineSpacing =
            rule === "exact"
              ? { rule: "exact", twips: line }
              : rule === "atleast"
                ? { rule: "atLeast", twips: line }
                : { rule: "auto", numerator: line, denominator: 240 };
      } else if (tag.local === "ind") {
        const left = Number(attr(tag, "left") ?? attr(tag, "start"));
        const right = Number(attr(tag, "right") ?? attr(tag, "end"));
        const firstLine = Number(attr(tag, "firstLine"));
        const hanging = Number(attr(tag, "hanging"));
        if (Number.isFinite(left) && left >= 0)
          current.values.leftIndentTwips = left;
        if (Number.isFinite(right) && right >= 0)
          current.values.rightIndentTwips = right;
        if (Number.isFinite(firstLine) && firstLine >= 0)
          current.values.firstLineIndentTwips = firstLine;
        if (Number.isFinite(hanging) && hanging >= 0)
          current.values.hangingIndentTwips = hanging;
      }
    },
    (tag) => {
      if (tag.local === "style" && current) {
        styles.set(current.id, current);
        current = null;
      }
    },
  );
  return styles;
};

const styleIdFor = (
  styles: ReadonlyMap<string, StylePatch>,
  candidates: readonly string[],
): string | null => {
  for (const candidate of candidates) {
    const exact = [...styles.values()].find(
      (style) =>
        style.id.toLowerCase() === candidate.toLowerCase() ||
        style.name?.toLowerCase() === candidate.toLowerCase(),
    );
    if (exact) return exact.id;
  }
  return null;
};

const inspectStyles = (
  styles: ReadonlyMap<string, StylePatch>,
  fallback: LayoutProfile,
  warnings: Diagnostic[],
): DocxTemplateInspection["styles"] => {
  const resolving = new Set<string>();
  const resolved = new Map<string, DocxTemplateInspection["styles"]["body"]>();
  const missingParents = new Set<string>();
  const cycles = new Set<string>();
  const resolve = (
    id: string | null,
    fallbackStyleValue: TextStyle,
    fallbackFamily: string,
    fallbackId: string,
  ): DocxTemplateInspection["styles"]["body"] => {
    if (!id)
      return fallbackStyle(fallbackStyleValue, fallbackFamily, fallbackId);
    const cached = resolved.get(id);
    if (cached) return cached;
    const style = styles.get(id);
    if (!style)
      return fallbackStyle(fallbackStyleValue, fallbackFamily, fallbackId);
    if (resolving.has(id)) {
      if (!cycles.has(id)) {
        cycles.add(id);
        warnings.push({
          code: "DOCX_STYLE_CYCLE",
          severity: "warning",
          message: `Template style inheritance cycle includes ${id}.`,
        });
      }
      return fallbackStyle(fallbackStyleValue, fallbackFamily, fallbackId);
    }
    resolving.add(id);
    let parent: DocxTemplateInspection["styles"]["body"] = fallbackStyle(
      fallbackStyleValue,
      fallbackFamily,
      fallbackId,
    );
    if (style.basedOn) {
      if (!styles.has(style.basedOn) && !missingParents.has(style.basedOn)) {
        missingParents.add(style.basedOn);
        warnings.push({
          code: "DOCX_STYLE_PARENT_MISSING",
          severity: "warning",
          message: `Template style ${id} references missing parent ${style.basedOn}.`,
        });
      } else
        parent = resolve(
          style.basedOn,
          fallbackStyleValue,
          fallbackFamily,
          fallbackId,
        );
    }
    resolving.delete(id);
    const value: DocxTemplateInspection["styles"]["body"] = {
      styleId: style.id,
      name: style.name,
      resolved: {
        ...parent.resolved,
        ...style.values,
        lineSpacing: style.values.lineSpacing ?? parent.resolved.lineSpacing,
      },
      requestedFontFamily: style.fontFamily ?? parent.requestedFontFamily,
      provenance: {
        ...parent.provenance,
        "": "template",
      },
    };
    resolved.set(id, value);
    return value;
  };
  const body = resolve(
    styleIdFor(styles, ["Normal"]),
    fallback.body,
    fallback.requestedFontFamily,
    "Normal",
  );
  const heading = (level: 1 | 2 | 3 | 4 | 5 | 6) =>
    resolve(
      styleIdFor(styles, [`Heading${level}`, `Heading ${level}`]),
      fallback.headings[String(level) as "1" | "2" | "3" | "4" | "5" | "6"],
      body.requestedFontFamily,
      `Heading${level}`,
    );
  return {
    body,
    headings: {
      "1": heading(1),
      "2": heading(2),
      "3": heading(3),
      "4": heading(4),
      "5": heading(5),
      "6": heading(6),
    },
    quote: resolve(
      styleIdFor(styles, ["Quote", "Intense Quote"]),
      fallback.blockquote,
      body.requestedFontFamily,
      "Quote",
    ),
    list: resolve(
      styleIdFor(styles, ["ListParagraph", "List Paragraph"]),
      fallback.list,
      body.requestedFontFamily,
      "ListParagraph",
    ),
    footnote: resolve(
      styleIdFor(styles, ["FootnoteText", "Footnote Text"]),
      fallback.footnote,
      body.requestedFontFamily,
      "FootnoteText",
    ),
    footnoteReference: (() => {
      const id = styleIdFor(styles, [
        "FootnoteReference",
        "Footnote Reference",
      ]);
      return id
        ? resolve(
            id,
            fallback.footnote,
            body.requestedFontFamily,
            "FootnoteReference",
          )
        : null;
    })(),
  };
};

const normalizedInstruction = (value: string): string =>
  value.replace(/\s+/g, " ").trim();

const fieldKind = (instruction: string): string =>
  normalizedInstruction(instruction).split(/\s+/, 1)[0]?.toUpperCase() ??
  "UNKNOWN";

const fieldsForPart = (
  xml: string,
  partPath: string,
): readonly DocxTemplateInspection["fields"][number][] => {
  const fields: DocxTemplateInspection["fields"][number][] = [];
  let complex: string | null = null;
  let inInstruction = false;
  const add = (instruction: string | undefined): void => {
    const normalized = normalizedInstruction(instruction ?? "");
    if (!normalized) return;
    fields.push({
      partPath,
      instruction: normalized,
      kind: fieldKind(normalized),
    });
  };
  parse(
    xml,
    (tag) => {
      if (tag.local === "fldSimple") {
        add(attr(tag, "instr"));
        return;
      }
      if (tag.local === "fldChar") {
        const kind = attr(tag, "fldCharType");
        if (kind === "begin") complex = "";
        else if (kind === "end" && complex !== null) {
          add(complex);
          complex = null;
        }
        return;
      }
      if (tag.local === "instrText" && complex !== null) inInstruction = true;
    },
    (tag) => {
      if (tag.local === "instrText") inInstruction = false;
    },
    (text) => {
      if (complex !== null && inInstruction) complex += text;
    },
  );
  return fields;
};

const textForPart = (xml: string): string => {
  let inText = false;
  let value = "";
  parse(
    xml,
    (tag) => {
      if (tag.local === "t" || tag.local === "delText") inText = true;
    },
    (tag) => {
      if (tag.local === "t" || tag.local === "delText") inText = false;
    },
    (text) => {
      if (inText) value += text;
    },
  );
  return value;
};

const captionsForPart = (
  xml: string,
): readonly DocxTemplateInspection["captions"][number][] => {
  const captions: DocxTemplateInspection["captions"][number][] = [];
  let paragraph: {
    index: number;
    styleId: string | null;
    text: string;
    instruction: string;
  } | null = null;
  let paragraphIndex = 0;
  let inText = false;
  let inInstruction = false;
  parse(
    xml,
    (tag) => {
      if (tag.local === "p") {
        paragraph = {
          index: paragraphIndex++,
          styleId: null,
          text: "",
          instruction: "",
        };
        return;
      }
      if (!paragraph) return;
      if (tag.local === "pStyle") paragraph.styleId = attr(tag, "val") ?? null;
      else if (tag.local === "fldSimple")
        paragraph.instruction += ` ${attr(tag, "instr") ?? ""}`;
      else if (tag.local === "t") inText = true;
      else if (tag.local === "instrText") inInstruction = true;
    },
    (tag) => {
      if (tag.local === "t") inText = false;
      else if (tag.local === "instrText") inInstruction = false;
      else if (tag.local === "p" && paragraph) {
        const instruction = normalizedInstruction(paragraph.instruction);
        const sequence = /\bSEQ\s+([^\s\\]+)/i.exec(instruction)?.[1] ?? null;
        if (/caption/i.test(paragraph.styleId ?? "") || sequence)
          captions.push({
            paragraphIndex: paragraph.index,
            styleId: paragraph.styleId,
            text: paragraph.text,
            sequence,
          });
        paragraph = null;
      }
    },
    (text) => {
      if (!paragraph) return;
      if (inText) paragraph.text += text;
      if (inInstruction) paragraph.instruction += text;
    },
  );
  return captions;
};

const numberingForPart = (
  xml: string | undefined,
): DocxTemplateInspection["numbering"] => {
  if (!xml) return { partPath: null, abstractNumbers: [], instances: [] };
  const abstractNumbers: Array<{ id: string; levels: number }> = [];
  const instances: Array<{ id: string; abstractNumberId: string | null }> = [];
  let abstract: { id: string; levels: number } | null = null;
  let instance: { id: string; abstractNumberId: string | null } | null = null;
  parse(
    xml,
    (tag) => {
      if (tag.local === "abstractNum") {
        const id = attr(tag, "abstractNumId");
        if (!id)
          throw new AgentDocxError(
            "DOCX_INVALID",
            "Numbering abstract ID is missing",
          );
        abstract = { id, levels: 0 };
      } else if (tag.local === "lvl" && abstract) abstract.levels++;
      else if (tag.local === "num") {
        const id = attr(tag, "numId");
        if (!id)
          throw new AgentDocxError(
            "DOCX_INVALID",
            "Numbering instance ID is missing",
          );
        instance = { id, abstractNumberId: null };
      } else if (tag.local === "abstractNumId" && instance)
        instance.abstractNumberId = attr(tag, "val") ?? null;
    },
    (tag) => {
      if (tag.local === "abstractNum" && abstract) {
        abstractNumbers.push(abstract);
        abstract = null;
      } else if (tag.local === "num" && instance) {
        instances.push(instance);
        instance = null;
      }
    },
  );
  return {
    partPath: "word/numbering.xml",
    abstractNumbers: abstractNumbers.sort((left, right) =>
      codePointCompare(left.id, right.id),
    ),
    instances: instances.sort((left, right) =>
      codePointCompare(left.id, right.id),
    ),
  };
};

const unsupportedPartReason = (name: string): string | null => {
  if (/vbaProject|macro/i.test(name))
    return "Macros are not imported from templates.";
  if (/\/(?:embeddings|activeX)\//i.test(name))
    return "Embedded executable content is not imported from templates.";
  if (/oleObject|altChunk/i.test(name))
    return "OLE or alternate content is not imported from templates.";
  return null;
};
export async function inspectDocxTemplate(
  docx: Uint8Array,
  options: InspectTemplateOptions = {},
): Promise<DocxTemplateInspection> {
  const fallback: LayoutProfile =
    typeof options.fallbackProfile === "string"
      ? builtInProfiles[options.fallbackProfile]
      : (options.fallbackProfile ??
        builtInProfiles["us-district-conventional"]);
  const { parts, names } = await readParts(docx);
  const content = parts["[Content_Types].xml"];
  if (!content)
    throw new AgentDocxError("DOCX_INVALID", "Missing [Content_Types].xml");
  const contentXml = decodeXml(content);
  const macroEnabled = /macroEnabled/i.test(contentXml);
  const rootRelationships = relationships(
    parts[relationshipPartFor("")]
      ? decodeXml(parts[relationshipPartFor("")]!)
      : undefined,
    "",
  ).filter((relationship) => /officeDocument$/.test(relationship.type));
  if (rootRelationships.length > 1)
    throw new AgentDocxError(
      "DOCX_INVALID",
      "DOCX has more than one main-document relationship",
    );
  const mainPart =
    rootRelationships.length === 0
      ? "word/document.xml"
      : (() => {
          const relationship = rootRelationships[0]!;
          if (relationship.external)
            throw new AgentDocxError(
              "DOCX_UNSAFE",
              "DOCX main-document relationship must be internal",
            );
          return resolveOpcTarget("", relationship.target);
        })();
  const main = parts[mainPart];
  if (!main)
    throw new AgentDocxError(
      "DOCX_INVALID",
      `Missing main document part: ${mainPart}`,
    );
  const mainXml = decodeXml(main);
  const sectionDefinitions = sectionGeometry(mainXml, fallback.page);
  const selectedSection = sectionDefinitions.length - 1;
  const warnings: Diagnostic[] = [];
  if (
    sectionDefinitions.some(
      (section) =>
        JSON.stringify(section.page) !==
        JSON.stringify(sectionDefinitions[selectedSection]!.page),
    )
  )
    warnings.push({
      code: "DOCX_MULTIPLE_SECTION_GEOMETRIES",
      severity: "warning",
      message:
        "Template contains differing section geometries; the final section is selected.",
    });
  const mainRelationships = relationships(
    parts[relationshipPartFor(mainPart)]
      ? decodeXml(parts[relationshipPartFor(mainPart)]!)
      : undefined,
    mainPart,
  );
  const relationshipById = new Map(
    mainRelationships.map((relationship) => [relationship.id, relationship]),
  );
  const sections = sectionDefinitions.map((section, index) => {
    const headerFooterReferences = section.headerFooterReferences.map(
      (reference) => {
        const relationship = relationshipById.get(reference.relationshipId);
        if (!relationship) {
          warnings.push({
            code: "DOCX_UNSUPPORTED_FEATURE",
            severity: "warning",
            message: `Section ${index} references missing ${reference.kind} relationship ${reference.relationshipId}.`,
          });
          return { ...reference, partPath: null };
        }
        if (
          relationship.external ||
          !new RegExp(`/${reference.kind}$`).test(relationship.type)
        ) {
          warnings.push({
            code: "DOCX_UNSUPPORTED_FEATURE",
            severity: "warning",
            message: `Section ${index} ${reference.kind} relationship ${reference.relationshipId} is not a supported internal ${reference.kind} part.`,
          });
          return { ...reference, partPath: null };
        }
        const partPath = resolveOpcTarget(mainPart, relationship.target);
        if (!parts[partPath]) {
          warnings.push({
            code: "DOCX_UNSUPPORTED_FEATURE",
            severity: "warning",
            message: `Section ${index} ${reference.kind} part is missing: ${partPath}.`,
          });
          return { ...reference, partPath: null };
        }
        return { ...reference, partPath };
      },
    );
    return {
      index,
      page: section.page,
      sourcePart: mainPart,
      headerFooterReferences,
    };
  });
  const headerFooters = sections.flatMap((section) =>
    section.headerFooterReferences.map((reference) => {
      const xml =
        reference.partPath === null
          ? undefined
          : decodeXml(parts[reference.partPath]!);
      return {
        sectionIndex: section.index,
        ...reference,
        text: xml ? textForPart(xml) : "",
        fields:
          xml && reference.partPath
            ? fieldsForPart(xml, reference.partPath)
            : [],
      };
    }),
  );
  const selected = sectionDefinitions[selectedSection]!.page;
  if (
    selected.widthTwips -
      selected.marginsTwips.left -
      selected.marginsTwips.right -
      selected.gutterTwips <=
      0 ||
    selected.heightTwips -
      selected.marginsTwips.top -
      selected.marginsTwips.bottom <=
      0
  )
    throw new AgentDocxError(
      "DOCX_INVALID",
      "Selected template section has non-positive usable area",
    );
  const theme = themeFonts(
    parts["word/theme/theme1.xml"]
      ? decodeXml(parts["word/theme/theme1.xml"]!)
      : undefined,
  );
  const parsedStyles = parseStyles(
    parts["word/styles.xml"] ? decodeXml(parts["word/styles.xml"]!) : undefined,
    theme,
  );
  const styles = inspectStyles(parsedStyles, fallback, warnings);
  const fieldPartPaths = Object.keys(parts)
    .filter((partPath) =>
      /^word\/(?:document|header\d+|footer\d+|footnotes)\.xml$/i.test(partPath),
    )
    .sort();
  const fields = fieldPartPaths.flatMap((partPath) =>
    fieldsForPart(decodeXml(parts[partPath]!), partPath),
  );
  const unsupportedParts = [
    ...names.flatMap((partPath) => {
      const reason = unsupportedPartReason(partPath);
      return reason ? [{ partPath, reason }] : [];
    }),
    ...Object.entries(parts)
      .filter(([partPath]) => partPath.endsWith(".rels"))
      .flatMap(([partPath, bytes]) =>
        relationships(decodeXml(bytes), sourcePartForRels(partPath))
          .filter((relationship) => relationship.external)
          .map((relationship) => ({
            partPath: `${partPath}#${relationship.id}`,
            reason: `External ${relationship.type} relationship is not copied from templates.`,
          })),
      ),
  ].sort((left, right) => codePointCompare(left.partPath, right.partPath));
  if (macroEnabled || unsupportedParts.length > 0)
    warnings.push({
      code: "DOCX_IGNORED_UNSAFE_PART",
      severity: "warning",
      message:
        "Unsafe, embedded, or external template parts are reported but never copied into generated documents.",
    });
  const families = new Map<string, string>();
  if (theme.major) families.set(theme.major, "word/theme/theme1.xml");
  if (theme.minor) families.set(theme.minor, "word/theme/theme1.xml");
  for (const style of [
    styles.body,
    ...Object.values(styles.headings),
    styles.quote,
    styles.list,
    styles.footnote,
    ...(styles.footnoteReference ? [styles.footnoteReference] : []),
  ])
    families.set(style.requestedFontFamily, "word/styles.xml");
  return {
    imported: {
      page: selected,
      requestedFontFamily: styles.body.requestedFontFamily,
      body: styles.body.resolved,
      headings: Object.fromEntries(
        Object.entries(styles.headings).map(([level, style]) => [
          level,
          style.resolved,
        ]),
      ),
      blockquote: styles.quote.resolved,
      list: styles.list.resolved,
      footnote: styles.footnote.resolved,
    },
    sections,
    selectedSection,
    styles,
    numbering: numberingForPart(
      parts["word/numbering.xml"]
        ? decodeXml(parts["word/numbering.xml"]!)
        : undefined,
    ),
    headerFooters,
    fields,
    captions: captionsForPart(mainXml),
    fonts: {
      theme,
      families: [...families.entries()]
        .sort(([left], [right]) => codePointCompare(left, right))
        .map(([family, sourcePart]) => ({ family, sourcePart })),
    },
    unsupportedParts,
    package: { sha256: sha256Hex(docx), mainPart, macroEnabled },
    warnings,
  };
}
