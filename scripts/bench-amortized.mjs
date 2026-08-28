#!/usr/bin/env node
// Bench: amortized (single-process) vs spawn-per-measure
// Ext4 Node24: library 15× ~500ms, CLI batch 15× ~700ms, CLI spawn 5× ~3s
import { performance } from "node:perf_hooks";
import { execSync } from "node:child_process";
import { measureMarkdown } from "../dist/index.js";

const mdBase =
  "# Test\n\n" +
  "Lorem ipsum dolor sit amet consectetur adipiscing elit sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ".repeat(
    30,
  );

async function libraryBench() {
  const opts = { profile: "us-district-conventional", lines: true };
  let t0 = performance.now();
  let r = await measureMarkdown(mdBase, opts);
  let t1 = performance.now();
  console.log(
    `library first (cold font parse): ${(t1 - t0).toFixed(1)} ms pages=${r.pageCount}`,
  );
  t0 = performance.now();
  for (let i = 0; i < 15; i++) {
    const variant = mdBase + ` Paragraph ${i} extra `.repeat(i * 3);
    r = await measureMarkdown(variant, opts);
  }
  t1 = performance.now();
  console.log(
    `library 15× amortized: ${(t1 - t0).toFixed(1)} ms avg=${((t1 - t0) / 15).toFixed(1)} ms`,
  );
  t0 = performance.now();
  for (let i = 0; i < 15; i++) {
    await measureMarkdown(mdBase + ` ${i} `.repeat(20), opts);
  }
  t1 = performance.now();
  console.log(
    `library 15× second loop (warm): ${(t1 - t0).toFixed(1)} ms avg=${((t1 - t0) / 15).toFixed(1)} ms`,
  );
}

function cliBench() {
  const requests = Array.from({ length: 15 }, (_, i) =>
    JSON.stringify({ id: i, markdown: mdBase + ` extra ${i} `.repeat(i * 5) }),
  ).join("\n");
  let t0 = performance.now();
  const out = execSync(
    `printf '%s' '${requests.replace(/'/g, "'\\''")}' | node ./dist/cli.js measure --batch --input-jsonl`,
    {
      encoding: "utf8",
    },
  );
  let t1 = performance.now();
  const lines = out.trim().split("\n").filter(Boolean).length;
  console.log(
    `CLI batch-jsonl 15× one spawn: ${(t1 - t0).toFixed(1)} ms results=${lines}`,
  );
  execSync(`printf '# Hello\\n\\nLorem\\n' > /tmp/bench-measure.md`);
  t0 = performance.now();
  for (let i = 0; i < 5; i++) {
    execSync(
      `node ./dist/cli.js measure /tmp/bench-measure.md --json > /dev/null`,
    );
  }
  t1 = performance.now();
  console.log(
    `CLI spawn-per-measure 5×: ${(t1 - t0).toFixed(1)} ms avg=${((t1 - t0) / 5).toFixed(1)} ms`,
  );
}

console.log("=== library ===");
await libraryBench();
console.log("\n=== CLI ===");
cliBench();
console.log(
  "\nDone — amortized should be <1s for 15, spawn ~6× slower on ext4 and ~300× on 9p.",
);
