#!/usr/bin/env node
import * as esbuild from "esbuild";
import { readFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";

const entry = "src/cli.ts";
const outfile = "dist/cli.js";

// Ensure esbuild can resolve .js imports to .ts sources via tsconfig resolution?
// We'll bundle from src directly; esbuild handles TS.

await esbuild.build({
  entryPoints: [entry],
  bundle: true,
  platform: "node",
  target: "node24",
  format: "esm",
  outfile,
  packages: "bundle", // bundle all deps
  sourcemap: true,
  minify: false,
  keepNames: true,
  // Provide `require` for bundled CJS dependencies that use require("fs") etc when output is ESM
  banner: {
    js: 'import { createRequire as __createRequire } from "node:module"; const require = __createRequire(import.meta.url);',
  },
  // Keep import.meta.url for asset/schema resolution
  // @resvg/resvg-js is native binding, externalize if present to avoid bundling .node
  external: ["@resvg/resvg-js", "@resvg/resvg-js-linux-x64-gnu"],
  loader: { ".json": "json" },
  logLevel: "info",
});

// Verify bundled CLI still has correct permissions and shebang handling
// esbuild banner adds shebang already, but ensure file starts correctly
const bundled = await readFile(fileURLToPath(new URL("../dist/cli.js", import.meta.url)), "utf8");
if (!bundled.startsWith("#!/usr/bin/env node")) {
  throw new Error("Bundled CLI missing shebang");
}
console.log(`Bundled CLI: ${outfile} (${(bundled.length / 1024).toFixed(1)} KB)`);
