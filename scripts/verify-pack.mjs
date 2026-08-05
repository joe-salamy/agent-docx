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
const packageJson = JSON.parse(
  await readFile(new URL("../package.json", import.meta.url), "utf8"),
);
const schemaNames = Object.keys(packageJson.exports)
  .filter((key) => key.endsWith(".schema.json"))
  .map((key) => key.slice(2));
if (schemaNames.length < 21)
  throw new Error(
    `Schema allowlist derivation found ${schemaNames.length} schemas; expected at least 21`,
  );
const fontPrefix = "assets/fonts/liberation-serif-2.1.5/";
const requiredRuleAssets = [
  "assets/rules/cand-civil-2026-05-01.txt",
  "assets/rules/frap-32-2024-12-01.txt",
];
const requiredAssets = [
  "assets/word/render.ps1",
  ...manifest.fonts.map(({ file }) => `${fontPrefix}${file}`),
  `${fontPrefix}manifest.json`,
  `${fontPrefix}OFL-1.1.txt`,
  ...requiredRuleAssets,
  ...schemaNames,
];
const npmCli = process.env.npm_execpath;
const usesNpmCli =
  npmCli !== undefined && /(?:^|[/\\])npm-cli\.js$/.test(npmCli);
const npmExecutable = usesNpmCli ? process.execPath : "npm";
const npmArguments = usesNpmCli ? [npmCli] : [];
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

const packDir = await mkdtemp(join(tmpdir(), "agent-docx-pack-"));
const installDir = await mkdtemp(join(tmpdir(), "agent-docx-install-"));
try {
  const packed = await run(
    npmExecutable,
    [
      ...npmArguments,
      "pack",
      "--ignore-scripts",
      "--json",
      "--pack-destination",
      packDir,
    ],
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
  const basenameOf = (path) => path.split("/").at(-1);
  const allowed = (path) =>
    /^(?:package\.json|README\.md|LICENSE|THIRD_PARTY_NOTICES\.txt|dist\/)/.test(
      path,
    ) ||
    (basenameOf(path).endsWith(".schema.json") &&
      schemaNames.includes(basenameOf(path))) ||
    requiredAssets.includes(path);
  const unexpected = paths.filter((path) => !allowed(path));
  if (unexpected.length) {
    throw new Error(
      `Archive contains unexpected files: ${unexpected.join(", ")}`,
    );
  }

  const missingAssets = requiredAssets.filter((path) => !paths.includes(path));
  if (missingAssets.length) {
    throw new Error(
      `Archive is missing required files: ${missingAssets.join(", ")}`,
    );
  }

  await writeFile(
    join(installDir, "package.json"),
    JSON.stringify({ name: "pack-smoke", private: true, type: "module" }),
  );
  const tgz = resolve(packDir, info.filename);
  await run(
    npmExecutable,
    [...npmArguments, "install", "--ignore-scripts", tgz],
    installDir,
  );
  const expected = JSON.stringify(expectedMetricFonts);
  await run(
    process.execPath,
    [
      "--input-type=module",
      "-e",
      `import {AgentDocxError,compileMarkdown,createProject,estimateMarkdown,inspectDocx,inspectDocxTemplate,measureMarkdown,openProject} from "agent-docx"; if (![AgentDocxError,compileMarkdown,createProject,estimateMarkdown,inspectDocx,inspectDocxTemplate,measureMarkdown,openProject].every(x=>typeof x==="function")) process.exit(1); const r=await estimateMarkdown("Smoke."); if(r.pageCount!==1) process.exit(2); const actual=r.profile.metricFonts.map(({role,metricsFamily,sha256})=>({role,metricsFamily,sha256})); if(JSON.stringify(actual)!==${JSON.stringify(expected)}) throw new Error(JSON.stringify(actual));`,
    ],
    installDir,
  );
  await run(
    process.execPath,
    [
      "--input-type=module",
      "-e",
      `import {readFile} from "node:fs/promises"; import {createRequire} from "node:module"; import {Ajv2020} from "ajv/dist/2020.js"; const require=createRequire(import.meta.url); const names=${JSON.stringify(schemaNames)}; const schemas=await Promise.all(names.map(async name=>JSON.parse(await readFile(require.resolve("agent-docx/"+name),"utf8")))); const ajv=new Ajv2020({strict:true,allowUnionTypes:true,formats:{date:true,"date-time":true,uri:true}}); for(const schema of schemas) ajv.addSchema(schema); for(const schema of schemas) ajv.getSchema(schema.$id);`,
    ],
    installDir,
  );
  await run(
    npmExecutable,
    [...npmArguments, "exec", "--", "agent-docx", "--version"],
    installDir,
  );
  console.log(`Verified ${info.filename} (${info.files.length} files)`);
} finally {
  await Promise.all([
    rm(packDir, { recursive: true, force: true }),
    rm(installDir, { recursive: true, force: true }),
  ]);
}
