import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import test from "node:test";
import { fileURLToPath } from "node:url";
import { root } from "./helpers.js";

const script = fileURLToPath(
  new URL("../scripts/accuracy.mjs", import.meta.url),
);

test("accuracy fails when the requested renderer is unavailable", async () => {
  const child = spawn(process.execPath, [script, "--renderer", "word"], {
    cwd: root,
    env: {
      ...process.env,
      AGENT_DOCX_TEST_WORD: "1",
      AGENT_DOCX_ACCURACY_WORD_PATH: "/definitely/missing/powershell",
    },
    stdio: ["ignore", "pipe", "pipe"],
  });
  let stdout = "";
  let stderr = "";
  child.stdout.setEncoding("utf8").on("data", (chunk) => (stdout += chunk));
  child.stderr.setEncoding("utf8").on("data", (chunk) => (stderr += chunk));
  const code = await new Promise((resolve, reject) => {
    child.once("error", reject);
    child.once("close", resolve);
  });
  assert.notEqual(code, 0, stdout);
  assert.match(stderr, /requested word renderer failed/i);
});
