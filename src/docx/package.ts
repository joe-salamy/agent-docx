import { deflateRawSync } from "node:zlib";
import yauzl, { type Entry, type ZipFile } from "yauzl";
import { TextDecoder } from "node:util";
import { SaxesParser, type SaxesTagNS } from "saxes";
import { createHash } from "node:crypto";
import { AgentDocxError } from "../types.js";
import { assertSafePartPath } from "./opc.js";
import { codePointCompare } from "./helpers.js";

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

export type DocxParts = ReadonlyMap<string, Uint8Array>;

export const sha256Hex = (bytes: Uint8Array | string): `sha256:${string}` =>
  `sha256:${createHash("sha256").update(bytes).digest("hex")}`;

const zipOpen = (bytes: Uint8Array): Promise<ZipFile> => {
  const { promise, resolve, reject } = Promise.withResolvers<ZipFile>();
  yauzl.fromBuffer(
    Buffer.from(bytes),
    { lazyEntries: true, validateEntrySizes: true, decodeStrings: true },
    (error, zip) => (error ? reject(error) : resolve(zip!)),
  );
  return promise;
};

const isXmlPart = (name: string): boolean =>
  name.endsWith(".xml") ||
  name.endsWith(".rels") ||
  name === "[Content_Types].xml";

const crc32Table = (() => {
  const table = new Uint32Array(256);
  for (let index = 0; index < table.length; index++) {
    let value = index;
    for (let bit = 0; bit < 8; bit++)
      value = value & 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
    table[index] = value >>> 0;
  }
  return table;
})();

const crc32Update = (value: number, bytes: Uint8Array): number => {
  for (const byte of bytes)
    value = crc32Table[(value ^ byte) & 0xff]! ^ (value >>> 8);
  return value >>> 0;
};

const crc32 = (bytes: Uint8Array): number =>
  (crc32Update(0xffffffff, bytes) ^ 0xffffffff) >>> 0;

const streamEntry = (
  zip: ZipFile,
  entry: Entry,
  maximum: number,
): Promise<Uint8Array> => {
  const { promise, resolve, reject } = Promise.withResolvers<Uint8Array>();
  zip.openReadStream(entry, (error, stream) => {
    if (error || !stream) {
      reject(error ?? new Error("No entry stream"));
      return;
    }
    const chunks: Buffer[] = [];
    let size = 0;
    let checksum = 0xffffffff;
    stream.on("data", (chunk: Buffer) => {
      size += chunk.length;
      if (size > maximum)
        stream.destroy(
          new AgentDocxError(
            isXmlPart(entry.fileName) ? "DOCX_XML_LIMIT" : "DOCX_TOO_LARGE",
            `Consumed package part exceeds ${maximum} bytes`,
          ),
        );
      else {
        checksum = crc32Update(checksum, chunk);
        chunks.push(chunk);
      }
    });
    stream.on("error", reject);
    stream.on("end", () => {
      const actual = (checksum ^ 0xffffffff) >>> 0;
      if (actual !== entry.crc32) {
        reject(
          new AgentDocxError(
            "DOCX_INVALID",
            `ZIP CRC-32 mismatch for package part: ${entry.fileName}`,
          ),
        );
        return;
      }
      resolve(Buffer.concat(chunks));
    });
  });
  return promise;
};

/**
 * Reads a DOCX OPC package without exposing ZIP paths or unbounded XML to callers.
 * Callers receive only normalized, unique entry names and must still follow explicit
 * relationship graphs before treating a part as reachable.
 */
export const readDocxParts = async (
  bytes: Uint8Array,
  include: (name: string) => boolean = () => true,
): Promise<DocxParts> => {
  if (bytes.byteLength > DOCX_LIMITS.maxCompressedInput)
    throw new AgentDocxError(
      "DOCX_TOO_LARGE",
      `DOCX exceeds ${DOCX_LIMITS.maxCompressedInput / (1024 * 1024)} MiB compressed input limit`,
    );
  const zip = await zipOpen(bytes);
  const parts = new Map<string, Uint8Array>();
  const names = new Set<string>();
  let entries = 0;
  let total = 0;
  let compressed = 0;
  let xmlTotal = 0;
  try {
    const { promise, resolve, reject } = Promise.withResolvers<void>();
    zip.on("error", reject);
    zip.on("end", resolve);
    zip.on("entry", (entry: Entry) => {
      void (async () => {
        try {
          entries++;
          if (entries > DOCX_LIMITS.maxEntries)
            throw new AgentDocxError(
              "DOCX_TOO_LARGE",
              `DOCX exceeds ${DOCX_LIMITS.maxEntries} entries`,
            );
          const name = entry.fileName;
          if (name.length > DOCX_LIMITS.maxPartNameUnits || names.has(name))
            throw new AgentDocxError(
              "DOCX_UNSAFE",
              `Unsafe or duplicate package path: ${name}`,
            );
          names.add(name);
          if (name.endsWith("/")) {
            assertSafePartPath(name.slice(0, -1));
            zip.readEntry();
            return;
          }
          assertSafePartPath(name);
          if (entry.compressionMethod !== 0 && entry.compressionMethod !== 8)
            throw new AgentDocxError(
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
            throw new AgentDocxError(
              "DOCX_TOO_LARGE",
              "DOCX expansion limit exceeded",
            );
          if (isXmlPart(name)) {
            xmlTotal += entry.uncompressedSize;
            if (xmlTotal > DOCX_LIMITS.maxXmlTotal)
              throw new AgentDocxError(
                "DOCX_XML_LIMIT",
                `Consumed XML exceeds ${DOCX_LIMITS.maxXmlTotal / (1024 * 1024)} MiB`,
              );
          }
          // Excluded entries are intentionally not consumed or CRC-verified.
          if (include(name))
            parts.set(
              name,
              await streamEntry(
                zip,
                entry,
                isXmlPart(name)
                  ? DOCX_LIMITS.maxXmlPart
                  : DOCX_LIMITS.maxUncompressedTotal,
              ),
            );
          zip.readEntry();
        } catch (error) {
          reject(error);
        }
      })();
    });
    zip.readEntry();
    await promise;
    if (total / Math.max(1, compressed) > DOCX_LIMITS.maxPackageRatio)
      throw new AgentDocxError(
        "DOCX_TOO_LARGE",
        `DOCX package compression ratio exceeds ${DOCX_LIMITS.maxPackageRatio}:1`,
      );
    return parts;
  } catch (error) {
    throw error instanceof AgentDocxError
      ? error
      : new AgentDocxError(
          "DOCX_INVALID",
          error instanceof Error ? error.message : String(error),
        );
  } finally {
    zip.close();
  }
};

export const decodeDocxXml = (bytes: Uint8Array): string => {
  let encoding = "utf-8";
  let offset = 0;
  if (bytes[0] === 0xff && bytes[1] === 0xfe) {
    encoding = "utf-16le";
    offset = 2;
  } else if (bytes[0] === 0xfe && bytes[1] === 0xff) {
    const swapped = new Uint8Array(bytes.length - 2);
    for (let index = 2; index < bytes.length; index += 2) {
      swapped[index - 2] = bytes[index + 1] ?? 0;
      swapped[index - 1] = bytes[index] ?? 0;
    }
    return new TextDecoder("utf-16le", { fatal: true }).decode(swapped);
  }
  try {
    return new TextDecoder(encoding, { fatal: true }).decode(
      bytes.subarray(offset),
    );
  } catch {
    throw new AgentDocxError("DOCX_INVALID", "XML is not valid UTF-8/UTF-16");
  }
};

export const parseDocxXml = (
  xml: string,
  onOpen: (tag: SaxesTagNS) => void,
  onClose?: (tag: SaxesTagNS) => void,
  onText?: (text: string) => void,
): void => {
  if (/<!DOCTYPE|<!ENTITY|<\?[^x]/i.test(xml))
    throw new AgentDocxError(
      "DOCX_UNSAFE",
      "DTD, entity, or non-XML processing instruction is forbidden",
    );
  let depth = 0;
  let elements = 0;
  let text = 0;
  const parser = new SaxesParser({ xmlns: true });
  parser.on("opentag", (tag) => {
    depth++;
    elements++;
    if (
      depth > DOCX_LIMITS.maxDepth ||
      elements > DOCX_LIMITS.maxElements ||
      Object.keys(tag.attributes).length > DOCX_LIMITS.maxAttributes
    )
      throw new AgentDocxError(
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
      throw new AgentDocxError("DOCX_XML_LIMIT", "XML text limit exceeded");
    onText?.(value);
  });
  try {
    parser.write(xml).close();
  } catch (error) {
    throw error instanceof AgentDocxError
      ? error
      : new AgentDocxError(
          "DOCX_INVALID",
          error instanceof Error ? error.message : String(error),
        );
  }
};

export const docxXmlAttribute = (
  tag: SaxesTagNS,
  name: string,
): string | undefined =>
  Object.values(tag.attributes).find((attribute) => attribute.local === name)
    ?.value;

export const resolveOpcTarget = (
  sourcePart: string,
  target: string,
): string => {
  if (target.startsWith("/") || target.includes("\\"))
    throw new AgentDocxError("DOCX_UNSAFE", "Unsafe relationship target");
  const source = sourcePart.split("/").slice(0, -1);
  const components = [...source, ...target.split("/")];
  const resolved: string[] = [];
  for (const component of components) {
    if (component === "" || component === ".") continue;
    if (component === "..") {
      if (resolved.length === 0)
        throw new AgentDocxError(
          "DOCX_UNSAFE",
          "Relationship escapes package root",
        );
      resolved.pop();
    } else resolved.push(component);
  }
  if (resolved.length === 0)
    throw new AgentDocxError("DOCX_UNSAFE", "Empty relationship target");
  return resolved.join("/");
};

const u16 = (value: number): Buffer => {
  const output = Buffer.allocUnsafe(2);
  output.writeUInt16LE(value, 0);
  return output;
};

const u32 = (value: number): Buffer => {
  const output = Buffer.allocUnsafe(4);
  output.writeUInt32LE(value >>> 0, 0);
  return output;
};

/**
 * Rebuilds a DOCX ZIP from trusted in-memory parts with sorted entries and fixed
 * ZIP timestamps. It intentionally supports only the ordinary ZIP subset emitted
 * by this package: UTF-8 names, deflate compression, and no ZIP64 extensions.
 */
export const repackDocxParts = (
  source: ReadonlyMap<string, Uint8Array>,
): Uint8Array => {
  const files = [...source.entries()]
    .sort(([left], [right]) => codePointCompare(left, right))
    .map(([name, bytes]) => {
      assertSafePartPath(name);
      const encodedName = Buffer.from(name, "utf8");
      const compressed = deflateRawSync(bytes, { level: 9 });
      if (
        encodedName.byteLength > 0xffff ||
        bytes.byteLength > 0xffffffff ||
        compressed.byteLength > 0xffffffff
      )
        throw new AgentDocxError(
          "DOCX_TOO_LARGE",
          "DOCX part cannot be represented in ZIP32",
        );
      return {
        name: encodedName,
        bytes,
        compressed,
        crc: crc32(bytes),
      };
    });
  const records: Buffer[] = [];
  const central: Buffer[] = [];
  let offset = 0;
  for (const file of files) {
    const local = Buffer.concat([
      u32(0x04034b50),
      u16(20),
      u16(0x0800),
      u16(8),
      u16(0),
      u16(33),
      u32(file.crc),
      u32(file.compressed.byteLength),
      u32(file.bytes.byteLength),
      u16(file.name.byteLength),
      u16(0),
      file.name,
      file.compressed,
    ]);
    records.push(local);
    central.push(
      Buffer.concat([
        u32(0x02014b50),
        u16(20),
        u16(20),
        u16(0x0800),
        u16(8),
        u16(0),
        u16(33),
        u32(file.crc),
        u32(file.compressed.byteLength),
        u32(file.bytes.byteLength),
        u16(file.name.byteLength),
        u16(0),
        u16(0),
        u16(0),
        u16(0),
        u32(0),
        u32(offset),
        file.name,
      ]),
    );
    offset += local.byteLength;
  }
  const centralBytes = Buffer.concat(central);
  if (
    files.length > 0xffff ||
    offset > 0xffffffff ||
    centralBytes.byteLength > 0xffffffff
  )
    throw new AgentDocxError(
      "DOCX_TOO_LARGE",
      "DOCX cannot be represented in ZIP32",
    );
  return Buffer.concat([
    ...records,
    centralBytes,
    u32(0x06054b50),
    u16(0),
    u16(0),
    u16(files.length),
    u16(files.length),
    u32(centralBytes.byteLength),
    u32(offset),
    u16(0),
  ]);
};
