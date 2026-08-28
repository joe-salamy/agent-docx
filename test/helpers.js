import { Ajv2020 } from "ajv/dist/2020.js";
import { spawn } from "node:child_process";
import { readFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { runCli } from "../dist/cli-run.js";

export const root = fileURLToPath(new URL("..", import.meta.url));
export const cli = fileURLToPath(new URL("../dist/cli.js", import.meta.url));
export const pkg = JSON.parse(
  await readFile(new URL("../package.json", import.meta.url), "utf8"),
);

export const metadata = {
  court: "United States District Court",
  jurisdiction: "Northern District of California",
  caseName: "Example v. Example",
  docketNumber: "3:26-cv-00001",
  documentTitle: "Motion",
  parties: [],
  counsel: [],
  certificates: [],
};

export const schemaNames = Object.keys(pkg.exports)
  .filter((key) => key.endsWith(".schema.json"))
  .map((key) => key.slice(2));
export const schemas = await Promise.all(
  schemaNames.map(async (name) =>
    JSON.parse(
      await readFile(new URL(`../schemas/${name}`, import.meta.url), "utf8"),
    ),
  ),
);
export const ajv = new Ajv2020({
  strict: true,
  allowUnionTypes: true,
  formats: { date: true, "date-time": true, uri: true },
});
for (const schema of schemas) ajv.addSchema(schema);
export const validatorFor = (id) => ajv.getSchema(id);
export const schemasByName = Object.fromEntries(
  schemaNames.map((name, index) => [name, schemas[index]]),
);
export const validatorForSchema = (name) =>
  ajv.getSchema(schemasByName[name]?.$id);

export function memoryRuntime(input = "", overrides = {}) {
  const stdout = [];
  const stderr = [];
  const runtime = {
    cwd: root,
    stdinIsTTY: false,
    version: pkg.version,
    readStdin: async () =>
      typeof input === "string" ? new TextEncoder().encode(input) : input,
    writeStdout: async (text) => void stdout.push(text),
    writeStderr: async (text) => void stderr.push(text),
    onceSignal() {},
    ...overrides,
  };
  return {
    runtime,
    stdout: () => stdout.join(""),
    stderr: () => stderr.join(""),
  };
}

export async function runInProcess(args, input = "", overrides = {}) {
  const capture = memoryRuntime(input, overrides);
  const code = await runCli(args, capture.runtime);
  return { code, stdout: capture.stdout(), stderr: capture.stderr() };
}

export function runSubprocess(args, input = "") {
  const { promise, resolve, reject } = Promise.withResolvers();
  const child = spawn(process.execPath, [cli, ...args], {
    cwd: root,
    stdio: ["pipe", "pipe", "pipe"],
  });
  let stdout = "";
  let stderr = "";
  child.stdout.setEncoding("utf8").on("data", (chunk) => (stdout += chunk));
  child.stderr.setEncoding("utf8").on("data", (chunk) => (stderr += chunk));
  child.once("error", reject);
  child.once("close", (code) => resolve({ code, stdout, stderr }));
  child.stdin.end(input);
  return promise;
}
