import type { SaxesTagNS } from "saxes";
import { AgentDocxError } from "../types.js";
import {
  docxXmlAttribute,
  parseDocxXml,
  resolveOpcTarget,
} from "./package.js";

export type Relationship = {
  id: string;
  type: string;
  target: string;
  external: boolean;
};

export const OOXML_NAMESPACES = Object.freeze({
  wordprocessingml:
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  relationships:
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  packageRelationships:
    "http://schemas.openxmlformats.org/package/2006/relationships",
} as const);

/**
 * Matches an OOXML element by local name AND expected namespace. Foreign
 * namespaces with colliding local names are treated as unknown, never parsed.
 */
export const isWordprocessingElement = (tag: SaxesTagNS): boolean =>
  tag.uri === OOXML_NAMESPACES.wordprocessingml;

export const isRelationshipsElement = (tag: SaxesTagNS): boolean =>
  tag.uri === OOXML_NAMESPACES.packageRelationships;

/**
 * One OPC relationship validator for every importer/inspector. Malformed
 * relationships XML is always `DOCX_INVALID`; semantic "valid but unsupported"
 * cases keep their own codes at the call sites.
 */
export const parseRelationships = (
  xml: string,
  sourcePart: string,
): readonly Relationship[] => {
  const found: Relationship[] = [];
  const ids = new Set<string>();
  parseDocxXml(xml, (tag) => {
    if (!isRelationshipsElement(tag) || tag.local !== "Relationship") return;
    const id = docxXmlAttribute(tag, "Id");
    const type = docxXmlAttribute(tag, "Type");
    const target = docxXmlAttribute(tag, "Target");
    const mode = docxXmlAttribute(tag, "TargetMode");
    if (!id || !type || !target || ids.has(id))
      throw new AgentDocxError("DOCX_INVALID", "Malformed relationship");
    ids.add(id);
    const external = mode === "External";
    if (mode !== undefined && mode !== "External")
      throw new AgentDocxError(
        "DOCX_INVALID",
        "Relationship target mode is invalid",
      );
    if (external) {
      if (!/\/hyperlink$/.test(type))
        throw new AgentDocxError(
          "DOCX_INVALID",
          "External non-hyperlink relationships are forbidden",
        );
      try {
        const url = new URL(target);
        if (!["http:", "https:", "mailto:"].includes(url.protocol))
          throw new Error("unsupported scheme");
      } catch {
        throw new AgentDocxError(
          "DOCX_INVALID",
          "External hyperlink has an unsupported target",
        );
      }
    } else resolveOpcTarget(sourcePart, target);
    found.push({ id, type, target, external });
  });
  return found;
};

/** Returns the `.rels` part path for a source part ("" for the package root). */
export const relationshipPartFor = (sourcePart: string): string => {
  if (sourcePart === "") return "_rels/.rels";
  const components = sourcePart.split("/");
  const name = components.pop();
  if (!name)
    throw new AgentDocxError("DOCX_INVALID", "Relationship source has no name");
  return [...components, "_rels", `${name}.rels`].join("/");
};

/** The strict OPC part-path policy: no empty, ".", or ".." components. */
export const assertSafePartPath = (name: string): void => {
  if (
    name.length === 0 ||
    name.includes("\\") ||
    name.startsWith("/") ||
    name
      .split("/")
      .some(
        (component) =>
          component === "" || component === "." || component === "..",
      )
  )
    throw new AgentDocxError("DOCX_UNSAFE", `Unsafe package path: ${name}`);
};
