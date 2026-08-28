import test from "node:test";
import assert from "node:assert/strict";
import { loadFonts } from "../dist/resolve.js";
import { readFile } from "node:fs/promises";

// Focused tests for parsed bundled-face single-flight caching (4d).
// Validates identity (same object) and output (same metrics/hash), not just timing.
// Does NOT cache whole deterministic results — only bundled font parsing.

test("bundled loadFonts is single-flight and memoizes by requestedFamily", async () => {
  // Sequential: same requestedFamily (case-insensitive) returns identical object
  const a = await loadFonts(undefined, "Liberation Serif");
  const b = await loadFonts(undefined, "liberation serif");
  const c = await loadFonts(undefined, "LIBERATION SERIF");
  assert.equal(a, b, "sequential same-family should be === (Map memo)");
  assert.equal(b, c, "case-insensitive key should be === ");
  // Different requestedFamily still shares parsed faces but different metrics wrapper
  // "Times New Roman" should substitute but share underlying font bytes/hashes
  const d = await loadFonts(undefined, "Times New Roman");
  assert.notEqual(a, d, "different requestedFamily should be different wrapper");
  assert.equal(a.regular.hash, d.regular.hash, "underlying bytes hash shared");
  assert.equal(a.regular.font, d.regular.font, "underlying font object shared via single-flight getBundledFaces");
  assert.equal(a.bold.font, d.bold.font);
  assert.equal(a.italic.font, d.italic.font);
  assert.equal(a.boldItalic.font, d.boldItalic.font);
});

test("bundled loadFonts concurrent calls are single-flight", async () => {
  // Concurrent: Promise.all should coalesce to one getBundledFaces parse
  const [p1, p2, p3] = await Promise.all([
    loadFonts(undefined, "Liberation Serif"),
    loadFonts(undefined, "Liberation Serif"),
    loadFonts(undefined, "Liberation Serif"),
  ]);
  assert.equal(p1, p2);
  assert.equal(p2, p3);
  // Also across different requestedFamily that share faces, the underlying faces promise is single-flight
  const [q1, q2] = await Promise.all([
    loadFonts(undefined, "Liberation Serif"),
    loadFonts(undefined, "Times New Roman"),
  ]);
  assert.equal(q1.regular.font, q2.regular.font, "concurrent across families still single-flight faces");
});

test("bundled loadFonts output matches manifest and is stable", async () => {
  const manifest = JSON.parse(
    await readFile(new URL("../assets/fonts/liberation-serif-2.1.5/manifest.json", import.meta.url), "utf8"),
  );
  const result = await loadFonts(undefined, "Liberation Serif");
  // Output: metricsFamily, sha256, and substitution flag
  assert.equal(result.family, "Liberation Serif");
  assert.deepEqual(
    result.metrics.map(({ role, metricsFamily, sha256 }) => ({ role, metricsFamily, sha256 })),
    manifest.fonts.map(({ sha256 }, i) => ({
      role: ["regular", "bold", "italic", "boldItalic"][i],
      metricsFamily: "Liberation Serif",
      sha256,
    })),
  );
  // No warnings for bundled
  assert.equal(result.warnings.length, 0);
  // Times substitution flag
  const times = await loadFonts(undefined, "Times New Roman");
  assert.equal(times.metrics[0].substitutedMetrics, true);
  assert.equal(times.metrics[0].metricsFamily, "Liberation Serif");
  assert.equal(times.family, "Liberation Serif");
});

test("custom fontSet bypasses bundled cache and validates", async () => {
  const bundled = await loadFonts(undefined, "Liberation Serif");
  const custom = await loadFonts(
    {
      family: "Liberation Serif",
      regular: bundled.regular.bytes,
      bold: bundled.bold.bytes,
      italic: bundled.italic.bytes,
      boldItalic: bundled.boldItalic.bytes,
    },
    "Liberation Serif",
  );
  // Custom path bypasses bundled memo (different object even though same bytes)
  assert.notEqual(bundled, custom);
  assert.equal(custom.family, "Liberation Serif");
  assert.equal(custom.regular.hash, bundled.regular.hash, "same bytes → same hash");
  // Custom metricsFamily reflects source.family, no substitution
  assert.equal(custom.metrics[0].metricsFamily, "Liberation Serif");
  assert.equal(custom.metrics[0].substitutedMetrics, false);
  // Mismatched family must reject
  await assert.rejects(
    () =>
      loadFonts(
        {
          family: "Custom",
          regular: bundled.regular.bytes,
          bold: bundled.bold.bytes,
          italic: bundled.italic.bytes,
          boldItalic: bundled.boldItalic.bytes,
        },
        "Custom",
      ),
    (e) => e.code === "INVALID_FONT",
  );
});
