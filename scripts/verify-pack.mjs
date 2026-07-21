import { spawn } from "node:child_process";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";

const fontDirectory = new URL(
  "../assets/fonts/liberation-serif-2.1.5/",
  import.meta.url,
);
const manifest = JSON.parse(
  await readFile(new URL("manifest.json", fontDirectory), "utf8"),
);
const roles = ["regular", "bold", "italic", "boldItalic"];
const expectedMetricFonts = manifest.fonts.map(({ sha256 }, index) => ({
  role: roles[index],
  metricsFamily: "Liberation Serif",
  sha256,
}));

function run(command, args, cwd) {
  const { promise, resolve: done, reject } = Promise.withResolvers();
  const child = spawn(command, args, {
    cwd,
    shell: false,
    stdio: ["ignore", "pipe", "pipe"],
  });
  let stdout = "";
  let stderr = "";
  child.stdout.setEncoding("utf8").on("data", (chunk) => (stdout += chunk));
  child.stderr.setEncoding("utf8").on("data", (chunk) => (stderr += chunk));
  child.once("error", reject);
  child.once("close", (code) =>
    code === 0
      ? done({ stdout, stderr })
      : reject(new Error(`${command} failed (${code}): ${stderr}`)),
  );
  return promise;
}

const packDir = await mkdtemp(join(tmpdir(), "md-page-count-pack-"));
const installDir = await mkdtemp(join(tmpdir(), "md-page-count-install-"));
try {
  const packed = await run(
    "npm",
    ["pack", "--json", "--pack-destination", packDir],
    process.cwd(),
  );
  const info = JSON.parse(packed.stdout)[0];
  if (!info?.filename || !Array.isArray(info.files)) {
    throw new Error("npm pack returned an invalid manifest");
  }

  const paths = info.files.map((file) => file.path);
  const forbidden = paths.filter(
    (path) =>
      /^(?:src|test|calibration|\.tmp)(?:\/|$)/.test(path) ||
      /(?:fixture|render\.docx|render\.pdf)/i.test(path) ||
      /^dist\/bundled-fonts(?:\.|$)/.test(path),
  );
  if (forbidden.length) {
    throw new Error(
      `Archive contains forbidden files: ${forbidden.join(", ")}`,
    );
  }
  const allowed =
    /^(?:package\.json|README\.md|LICENSE|THIRD_PARTY_NOTICES\.txt|config\.schema\.json|dist\/|assets\/)/;
  const unexpected = paths.filter((path) => !allowed.test(path));
  if (unexpected.length) {
    throw new Error(
      `Archive contains unexpected files: ${unexpected.join(", ")}`,
    );
  }

  const fontPrefix = "assets/fonts/liberation-serif-2.1.5/";
  const requiredAssets = [
    ...manifest.fonts.map(({ file }) => `${fontPrefix}${file}`),
    `${fontPrefix}manifest.json`,
    `${fontPrefix}OFL-1.1.txt`,
  ];
  const missingAssets = requiredAssets.filter((path) => !paths.includes(path));
  if (missingAssets.length) {
    throw new Error(
      `Archive is missing font assets: ${missingAssets.join(", ")}`,
    );
  }

  await writeFile(
    join(installDir, "package.json"),
    JSON.stringify({ name: "pack-smoke", private: true, type: "module" }),
  );
  const tgz = resolve(packDir, info.filename);
  await run("npm", ["install", "--ignore-scripts", tgz], installDir);
  const expected = JSON.stringify(expectedMetricFonts);
  await run(
    process.execPath,
    [
      "--input-type=module",
      "-e",
      `import {estimateMarkdown,measureMarkdown,inspectDocxTemplate} from "md-page-count"; if (![estimateMarkdown,measureMarkdown,inspectDocxTemplate].every(x=>typeof x==="function")) process.exit(1); const r=await estimateMarkdown("Smoke."); if(r.pageCount!==1) process.exit(2); const actual=r.profile.metricFonts.map(({role,metricsFamily,sha256})=>({role,metricsFamily,sha256})); if(JSON.stringify(actual)!==${JSON.stringify(expected)}) throw new Error(JSON.stringify(actual));`,
    ],
    installDir,
  );
  await run("npm", ["exec", "--", "md-page-count", "--version"], installDir);
  console.log(`Verified ${info.filename} (${info.files.length} files)`);
} finally {
  await Promise.all([
    rm(packDir, { recursive: true, force: true }),
    rm(installDir, { recursive: true, force: true }),
  ]);
}
