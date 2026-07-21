import yauzl, { type Entry, type ZipFile } from "yauzl";
import { createHash } from "node:crypto";
import { TextDecoder } from "node:util";
import { SaxesParser, type SaxesTagNS } from "saxes";
import { builtInProfiles } from "../profiles.js";
import {
  MdPageCountError,
  type Diagnostic,
  type DocxTemplateInspection,
  type InspectTemplateOptions,
  type LayoutProfile,
  type PageGeometry,
  type TextStyle,
} from "../types.js";
export const DOCX_LIMITS = Object.freeze({
  maxCompressedInput: 25 * 1024 * 1024,
  maxEntries: 512,
  maxPartNameUnits: 240,
  maxUncompressedTotal: 64 * 1024 * 1024,
  maxEntryRatio: 100,
  maxPackageRatio: 50,
  maxXmlPart: 4 * 1024 * 1024,
  maxXmlTotal: 12 * 1024 * 1024,
  maxDepth: 64,
  maxElements: 100000,
  maxAttributes: 128,
  maxText: 1024 * 1024,
});
type Parts = Record<string, Uint8Array>;
const sha = (bytes: Uint8Array) =>
  createHash("sha256").update(bytes).digest("hex");
function zipOpen(bytes: Uint8Array): Promise<ZipFile> {
  return new Promise((resolve, reject) =>
    yauzl.fromBuffer(
      Buffer.from(bytes),
      { lazyEntries: true, validateEntrySizes: true, decodeStrings: true },
      (error, zip) => (error ? reject(error) : resolve(zip!)),
    ),
  );
}
function streamEntry(zip: ZipFile, entry: Entry): Promise<Uint8Array> {
  return new Promise((resolve, reject) =>
    zip.openReadStream(entry, (error, stream) => {
      if (error || !stream) {
        reject(error ?? new Error("No entry stream"));
        return;
      }
      const chunks: Buffer[] = [];
      let size = 0;
      stream.on("data", (chunk: Buffer) => {
        size += chunk.length;
        if (size > DOCX_LIMITS.maxXmlPart)
          stream.destroy(
            new MdPageCountError(
              "DOCX_XML_LIMIT",
              "Consumed XML part exceeds 4 MiB",
            ),
          );
        else chunks.push(chunk);
      });
      stream.on("error", reject);
      stream.on("end", () => resolve(Buffer.concat(chunks)));
    }),
  );
}
async function readParts(bytes: Uint8Array): Promise<Parts> {
  if (bytes.byteLength > DOCX_LIMITS.maxCompressedInput)
    throw new MdPageCountError(
      "DOCX_TOO_LARGE",
      "DOCX exceeds 25 MiB compressed input limit",
    );
  const zip = await zipOpen(bytes);
  const parts: Parts = {};
  let entries = 0,
    total = 0,
    compressed = 0,
    xmlTotal = 0;
  try {
    await new Promise<void>((resolve, reject) => {
      zip.on("error", reject);
      zip.on("end", resolve);
      zip.on("entry", (entry: Entry) => {
        void (async () => {
          try {
            entries++;
            if (entries > DOCX_LIMITS.maxEntries)
              throw new MdPageCountError(
                "DOCX_TOO_LARGE",
                "DOCX exceeds 512 entries",
              );
            const name = entry.fileName;
            if (
              name.length > DOCX_LIMITS.maxPartNameUnits ||
              name.includes("\\") ||
              name.startsWith("/") ||
              name.split("/").includes("..") ||
              Object.hasOwn(parts, name)
            )
              throw new MdPageCountError(
                "DOCX_UNSAFE",
                `Unsafe or duplicate package path: ${name}`,
              );
            if (/\/$/.test(name)) {
              zip.readEntry();
              return;
            }
            if (entry.compressionMethod !== 0 && entry.compressionMethod !== 8)
              throw new MdPageCountError(
                "DOCX_UNSAFE",
                "Unsupported ZIP compression method",
              );
            total += entry.uncompressedSize;
            compressed += entry.compressedSize;
            if (
              total > DOCX_LIMITS.maxUncompressedTotal ||
              entry.uncompressedSize / Math.max(1, entry.compressedSize) >
                DOCX_LIMITS.maxEntryRatio
            )
              throw new MdPageCountError(
                "DOCX_TOO_LARGE",
                "DOCX expansion limit exceeded",
              );
            const relevant =
              name === "[Content_Types].xml" ||
              name.endsWith(".rels") ||
              name.endsWith("document.xml") ||
              name.endsWith("styles.xml") ||
              name.endsWith("theme1.xml") ||
              name.endsWith("settings.xml");
            if (relevant) {
              xmlTotal += entry.uncompressedSize;
              if (xmlTotal > DOCX_LIMITS.maxXmlTotal)
                throw new MdPageCountError(
                  "DOCX_XML_LIMIT",
                  "Consumed XML exceeds 12 MiB",
                );
              parts[name] = await streamEntry(zip, entry);
            }
            zip.readEntry();
          } catch (error) {
            reject(error);
          }
        })();
      });
      zip.readEntry();
    });
    if (total / Math.max(1, compressed) > DOCX_LIMITS.maxPackageRatio)
      throw new MdPageCountError(
        "DOCX_TOO_LARGE",
        "DOCX package compression ratio exceeds 50:1",
      );
    return parts;
  } catch (error) {
    throw error instanceof MdPageCountError
      ? error
      : new MdPageCountError(
          "DOCX_INVALID",
          error instanceof Error ? error.message : String(error),
        );
  } finally {
    zip.close();
  }
}
function decodeXml(bytes: Uint8Array): string {
  let encoding = "utf-8",
    offset = 0;
  if (bytes[0] === 0xff && bytes[1] === 0xfe) {
    encoding = "utf-16le";
    offset = 2;
  } else if (bytes[0] === 0xfe && bytes[1] === 0xff) {
    const swapped = new Uint8Array(bytes.length - 2);
    for (let i = 2; i < bytes.length; i += 2) {
      swapped[i - 2] = bytes[i + 1] ?? 0;
      swapped[i - 1] = bytes[i] ?? 0;
    }
    return new TextDecoder("utf-16le", { fatal: true }).decode(swapped);
  }
  try {
    return new TextDecoder(encoding, { fatal: true }).decode(
      bytes.subarray(offset),
    );
  } catch {
    throw new MdPageCountError("DOCX_INVALID", "XML is not valid UTF-8/UTF-16");
  }
}
function parse(
  xml: string,
  onOpen: (tag: SaxesTagNS) => void,
  onClose?: (tag: SaxesTagNS) => void,
) {
  if (/<!DOCTYPE|<!ENTITY|<\?[^x]/i.test(xml))
    throw new MdPageCountError(
      "DOCX_UNSAFE",
      "DTD, entity, or non-XML processing instruction is forbidden",
    );
  let depth = 0,
    elements = 0,
    text = 0;
  const parser = new SaxesParser({ xmlns: true });
  parser.on("opentag", (tag) => {
    depth++;
    elements++;
    if (
      depth > DOCX_LIMITS.maxDepth ||
      elements > DOCX_LIMITS.maxElements ||
      Object.keys(tag.attributes).length > DOCX_LIMITS.maxAttributes
    )
      throw new MdPageCountError(
        "DOCX_XML_LIMIT",
        "XML structural limit exceeded",
      );
    onOpen(tag);
  });
  parser.on("closetag", (tag) => {
    onClose?.(tag);
    depth--;
  });
  parser.on("text", (value) => {
    text += value.length;
    if (text > DOCX_LIMITS.maxText)
      throw new MdPageCountError("DOCX_XML_LIMIT", "XML text limit exceeded");
  });
  try {
    parser.write(xml).close();
  } catch (error) {
    throw error instanceof MdPageCountError
      ? error
      : new MdPageCountError(
          "DOCX_INVALID",
          error instanceof Error ? error.message : String(error),
        );
  }
}
const attr = (tag: SaxesTagNS, name: string) =>
  Object.values(tag.attributes).find((a) => a.local === name)?.value;
function sectionGeometry(
  xml: string,
  sourcePart: string,
  fallback: PageGeometry,
): PageGeometry[] {
  const sections: PageGeometry[] = [];
  let page: PageGeometry | null = null;
  parse(
    xml,
    (tag) => {
      if (tag.local === "sectPr") page = structuredClone(fallback);
      else if (page && tag.local === "pgSz") {
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
      }
    },
    (tag) => {
      if (tag.local === "sectPr" && page) {
        sections.push(page);
        page = null;
      }
    },
  );
  if (sections.length === 0) sections.push(structuredClone(fallback));
  void sourcePart;
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
function makeHeadings(
  fallback: LayoutProfile,
): DocxTemplateInspection["styles"]["headings"] {
  return {
    "1": fallbackStyle(
      fallback.headings["1"],
      fallback.requestedFontFamily,
      "Heading1",
    ),
    "2": fallbackStyle(
      fallback.headings["2"],
      fallback.requestedFontFamily,
      "Heading2",
    ),
    "3": fallbackStyle(
      fallback.headings["3"],
      fallback.requestedFontFamily,
      "Heading3",
    ),
    "4": fallbackStyle(
      fallback.headings["4"],
      fallback.requestedFontFamily,
      "Heading4",
    ),
    "5": fallbackStyle(
      fallback.headings["5"],
      fallback.requestedFontFamily,
      "Heading5",
    ),
    "6": fallbackStyle(
      fallback.headings["6"],
      fallback.requestedFontFamily,
      "Heading6",
    ),
  };
}
export async function inspectDocxTemplate(
  docx: Uint8Array,
  options: InspectTemplateOptions = {},
): Promise<DocxTemplateInspection> {
  const fallback: LayoutProfile =
    typeof options.fallbackProfile === "string"
      ? builtInProfiles[options.fallbackProfile]
      : (options.fallbackProfile ??
        builtInProfiles["us-district-conventional"]);
  const parts = await readParts(docx);
  const content = parts["[Content_Types].xml"];
  if (!content)
    throw new MdPageCountError("DOCX_INVALID", "Missing [Content_Types].xml");
  const contentXml = decodeXml(content);
  const macroEnabled = /macroEnabled/i.test(contentXml);
  let mainPart = "word/document.xml";
  const rootRels = parts["_rels/.rels"];
  if (rootRels) {
    parse(decodeXml(rootRels), (tag) => {
      if (
        tag.local === "Relationship" &&
        /officeDocument$/.test(attr(tag, "Type") ?? "")
      ) {
        const target = attr(tag, "Target");
        if (target) mainPart = target.replace(/^\//, "");
      }
    });
  }
  const main = parts[mainPart];
  if (!main)
    throw new MdPageCountError(
      "DOCX_INVALID",
      `Missing main document part: ${mainPart}`,
    );
  const geometries = sectionGeometry(decodeXml(main), mainPart, fallback.page);
  const selectedSection = geometries.length - 1;
  const warnings: Diagnostic[] = [];
  if (
    geometries.some(
      (g) => JSON.stringify(g) !== JSON.stringify(geometries[selectedSection]),
    )
  )
    warnings.push({
      code: "DOCX_MULTIPLE_SECTION_GEOMETRIES",
      severity: "warning",
      message:
        "Template contains differing section geometries; the final section is selected.",
    });
  if (macroEnabled)
    warnings.push({
      code: "DOCX_IGNORED_UNSAFE_PART",
      severity: "warning",
      message:
        "Macro content is ignored; generated renderer documents are non-macro.",
    });
  const selected = geometries[selectedSection]!;
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
    throw new MdPageCountError(
      "DOCX_INVALID",
      "Selected template section has non-positive usable area",
    );
  const body = fallbackStyle(
    fallback.body,
    fallback.requestedFontFamily,
    "Normal",
  );
  const headingEntries = makeHeadings(fallback);
  return {
    imported: { page: selected },
    sections: geometries.map((page, index) => ({
      index,
      page,
      sourcePart: mainPart,
    })),
    selectedSection,
    styles: {
      body,
      headings: headingEntries,
      quote: fallbackStyle(
        fallback.blockquote,
        fallback.requestedFontFamily,
        "Quote",
      ),
      footnote: fallbackStyle(
        fallback.footnote,
        fallback.requestedFontFamily,
        "FootnoteText",
      ),
      footnoteReference: null,
    },
    package: { sha256: sha(docx), mainPart, macroEnabled },
    warnings,
  };
}
