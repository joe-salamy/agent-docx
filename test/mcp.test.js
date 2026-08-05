import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import test from "node:test";
import { fileURLToPath } from "node:url";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { Ajv2020 } from "ajv/dist/2020.js";
import { metadata } from "./helpers.js";

const cli = fileURLToPath(new URL("../dist/cli.js", import.meta.url));
const supportedProtocolVersions = new Set([
  "2024-11-05",
  "2025-03-26",
  "2025-06-18",
]);

const waitForClose = (child, timeoutMs) =>
  new Promise((resolve, reject) => {
    let settled = false;
    const timer = setTimeout(() => {
      if (settled) return;
      settled = true;
      reject(new Error("MCP subprocess did not close before timeout"));
    }, timeoutMs);
    child.once("close", (code, signal) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      resolve({ code, signal });
    });
  });

const spawnMcp = (cwd) => {
  const child = spawn(process.execPath, [cli, "mcp"], {
    cwd,
    stdio: ["pipe", "pipe", "pipe"],
  });
  let buffer = "";
  let closed = null;
  const lines = [];
  const waiters = [];

  const rejectWaiters = (error) => {
    while (waiters.length > 0) waiters.shift().reject(error);
  };
  child.stdout.setEncoding("utf8");
  child.stdout.on("data", (chunk) => {
    buffer += chunk;
    for (;;) {
      const newline = buffer.indexOf("\n");
      if (newline < 0) break;
      const line = buffer.slice(0, newline);
      buffer = buffer.slice(newline + 1);
      if (line.trim() === "") continue;
      const waiter = waiters.shift();
      if (waiter) waiter.resolve(line);
      else lines.push(line);
    }
  });
  child.stderr.setEncoding("utf8");
  let stderr = "";
  child.stderr.on("data", (chunk) => (stderr += chunk));
  child.once("error", (error) => {
    closed = error;
    rejectWaiters(error);
  });
  child.once("close", (code, signal) => {
    closed = new Error(
      `MCP subprocess closed before a response (code=${code}, signal=${signal}, stderr=${stderr})`,
    );
    rejectWaiters(closed);
  });

  return {
    child,
    stderr: () => stderr,
    send(message) {
      child.stdin.write(`${JSON.stringify(message)}\n`);
    },
    nextLine(timeoutMs = 30000) {
      if (lines.length > 0) return Promise.resolve(lines.shift());
      if (closed) return Promise.reject(closed);
      return new Promise((resolve, reject) => {
        const waiter = {
          resolve: (line) => {
            clearTimeout(timer);
            resolve(line);
          },
          reject: (error) => {
            clearTimeout(timer);
            reject(error);
          },
        };
        const timer = setTimeout(() => {
          const index = waiters.indexOf(waiter);
          if (index >= 0) waiters.splice(index, 1);
          reject(new Error("Timed out waiting for MCP response"));
        }, timeoutMs);
        waiters.push(waiter);
      });
    },
    async close() {
      if (!child.stdin.destroyed) child.stdin.end();
      try {
        return await waitForClose(child, 15000);
      } catch (error) {
        child.kill("SIGKILL");
        await waitForClose(child, 15000).catch(() => undefined);
        throw error;
      }
    },
  };
};

const call = async (session, id, method, params) => {
  session.send({
    jsonrpc: "2.0",
    id,
    method,
    ...(params === undefined ? {} : { params }),
  });
  return JSON.parse(await session.nextLine());
};

test("MCP stdio serves agent tools and dispatches a project workflow", async () => {
  const directory = await mkdtemp(join(tmpdir(), "agent-docx-mcp-"));
  const session = spawnMcp(directory);
  try {
    const initialized = await call(session, "initialize", "initialize", {
      protocolVersion: "unsupported-version",
    });
    assert.equal(initialized.jsonrpc, "2.0");
    assert.ok(
      supportedProtocolVersions.has(initialized.result.protocolVersion),
    );
    assert.equal(initialized.result.capabilities.tools.listChanged, false);

    session.send({
      jsonrpc: "2.0",
      method: "notifications/initialized",
      params: {},
    });
    const listed = await call(session, "list", "tools/list", {});
    const names = listed.result.tools.map((tool) => tool.name);
    assert.ok(names.includes("project.init"));
    assert.ok(names.includes("document.validate"));
    assert.ok(names.includes("docx.export"));
    for (const tool of listed.result.tools) {
      const ajv = new Ajv2020({ strict: false });
      assert.doesNotThrow(
        () => ajv.compile(tool.inputSchema),
        `tool ${tool.name} inputSchema must compile with no dangling refs`,
      );
    }
    const initSchema = listed.result.tools.find(
      (tool) => tool.name === "project.init",
    ).inputSchema;
    assert.ok(
      ["documentId", "source", "profile", "metadata"].every((key) =>
        initSchema.required.includes(key),
      ),
      "project.init inputSchema should carry action-specific required params",
    );
    assert.equal(typeof initSchema.properties.documentId.$ref, "string");
    assert.equal(typeof initSchema.properties.project.type, "string");

    const badProject = await call(session, "bad-project", "tools/call", {
      name: "project.init",
      arguments: { project: 42 },
    });
    assert.equal(badProject.result, undefined);
    assert.equal(badProject.error.code, -32602);
    assert.equal(badProject.error.message, "project must be a non-empty string");

    const initializedProject = await call(session, "project", "tools/call", {
      name: "project.init",
      arguments: {
        project: "agent-docx.json",
        documentId: "motion",
        source: "motion.md",
        createSource: true,
        profile: "us-district-conventional",
        metadata,
      },
    });
    assert.equal(initializedProject.result.isError, false);
    assert.ok(
      Array.isArray(initializedProject.result.structuredContent.documents),
    );

    const document = await call(session, "document", "tools/call", {
      name: "document.get",
      arguments: { project: "agent-docx.json", documentId: "motion" },
    });
    assert.equal(document.result.isError, false);
    assert.ok(document.result.structuredContent.document);

    const validation = await call(session, "validation", "tools/call", {
      name: "document.validate",
      arguments: { project: "agent-docx.json", documentId: "motion" },
    });
    assert.equal(validation.result.isError, false);
    assert.ok(Array.isArray(validation.result.structuredContent.findings));

    const knownActionFailure = await call(
      session,
      "known-action-failure",
      "tools/call",
      {
        name: "document.validate",
        arguments: { project: "agent-docx.json" },
      },
    );
    assert.equal(knownActionFailure.result.isError, true);
    assert.equal(knownActionFailure.error, undefined);

    const unknownTool = await call(session, "unknown-tool", "tools/call", {
      name: "not.an.action",
      arguments: {},
    });
    assert.equal(unknownTool.result, undefined);
    assert.equal(unknownTool.error.code, -32602);

    const unknownMethod = await call(
      session,
      "unknown-method",
      "tools/unknown",
      {},
    );
    assert.equal(unknownMethod.error.code, -32601);

    session.send([
      { jsonrpc: "2.0", id: "first", method: "ping" },
      { jsonrpc: "2.0", id: "second", method: "ping" },
    ]);
    const batch = JSON.parse(await session.nextLine());
    assert.equal(batch.error.code, -32600);
  } finally {
    await session.close().catch((error) => {
      assert.fail(`${error.message}\n${session.stderr()}`);
    });
    await rm(directory, { recursive: true, force: true });
  }
});

test("MCP error responses carry structured details", async () => {
  const directory = await mkdtemp(join(tmpdir(), "agent-docx-mcp-details-"));
  await writeFile(join(directory, "agent-docx.json"), "{not json");
  const session = spawnMcp(directory);
  try {
    await call(session, "initialize", "initialize", {
      protocolVersion: "2025-06-18",
    });
    session.send({
      jsonrpc: "2.0",
      method: "notifications/initialized",
      params: {},
    });
    const failed = await call(session, "details", "tools/call", {
      name: "document.get",
      arguments: { project: "agent-docx.json", documentId: "motion" },
    });
    assert.equal(failed.result.isError, true);
    assert.equal(failed.result.structuredContent.code, "PROJECT_INVALID");
    assert.ok(
      failed.result.structuredContent.details !== undefined,
      "structuredContent should include details",
    );
  } finally {
    await session.close().catch((error) => {
      assert.fail(`${error.message}\n${session.stderr()}`);
    });
    await rm(directory, { recursive: true, force: true });
  }
});
