import { mkdir, readFile, writeFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { Resvg } from "@resvg/resvg-js";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const previewSource = resolve(root, "docs/assets/agent-docx-social-preview.svg");
const output = resolve(root, "docs/assets/agent-docx-social-preview.png");
const fontDirectory = resolve(root, "assets/fonts/liberation-serif-2.1.5");
const expectedWidth = 1280;
const expectedHeight = 640;
const maximumBytes = 1024 * 1024;
const pngSignature = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);

function assertPreviewPng(png) {
  const bytes = Buffer.from(png);

  if (!bytes.subarray(0, pngSignature.length).equals(pngSignature)) {
    throw new Error("Social preview renderer did not produce a PNG.");
  }

  if (
    bytes.length < 24 ||
    bytes.readUInt32BE(8) !== 13 ||
    bytes.toString("ascii", 12, 16) !== "IHDR"
  ) {
    throw new Error("Social preview PNG is missing a valid IHDR chunk.");
  }

  const width = bytes.readUInt32BE(16);
  const height = bytes.readUInt32BE(20);
  if (width !== expectedWidth || height !== expectedHeight) {
    throw new Error(
      `Social preview must be ${expectedWidth}×${expectedHeight}, received ${width}×${height}.`,
    );
  }

  if (bytes.length >= maximumBytes) {
    throw new Error(
      `Social preview must be smaller than 1 MiB, received ${bytes.length} bytes.`,
    );
  }
}

const svg = await readFile(previewSource, "utf8");
const png = new Resvg(svg, {
  font: {
    defaultFontFamily: "Liberation Serif",
    fontFiles: [
      resolve(fontDirectory, "LiberationSerif-Regular.ttf"),
      resolve(fontDirectory, "LiberationSerif-Bold.ttf"),
    ],
    loadSystemFonts: false,
  },
}).render().asPng();

assertPreviewPng(png);
await mkdir(dirname(output), { recursive: true });
await writeFile(output, png);
console.log("Wrote docs/assets/agent-docx-social-preview.png (1280×640)");
