import test from "node:test";
import assert from "node:assert/strict";
import { loadFonts } from "../dist/resolve.js";
import { readFile } from "node:fs/promises";

// Focused tests for parsed bundled-face single-flight caching (4d).
// Validates identity (same object) and output (same metrics/hash), not just timing.
// Does NOT cache whole deterministic results — only bundled font parsing.

test("bundled loadFonts is single-flight for faces but fresh wrapper per call", async () => {
  const a = await loadFonts(undefined, "Liberation Serif");
  const b = await loadFonts(undefined, "liberation serif");
  const c = await loadFonts(undefined, "LIBERATION SERIF");
  // Wrappers are fresh each call — requestedFamily must be exact, and metrics
  // array must not be shared mutable state exposed via public results
  assert.notEqual(
    a,
    b,
    "sequential calls must be fresh wrappers (requestedFamily isolation)",
  );
  assert.notEqual(b, c);
  assert.notEqual(a.metrics, b.metrics, "metrics array must not be shared");
  assert.notEqual(a.warnings, b.warnings, "warnings array must not be shared");
  assert.equal(a.metrics[0].requestedFamily, "Liberation Serif");
  assert.equal(b.metrics[0].requestedFamily, "liberation serif");
  assert.equal(c.metrics[0].requestedFamily, "LIBERATION SERIF");
  // Underlying parsed faces (bytes/hash/font) are single-flight shared
  assert.equal(a.regular.hash, b.regular.hash, "underlying bytes hash shared");
  assert.equal(
    a.regular.font,
    b.regular.font,
    "underlying font object shared via single-flight getBundledFaces",
  );
  assert.equal(a.bold.font, b.bold.font);
  assert.equal(a.italic.font, b.italic.font);
  assert.equal(a.boldItalic.font, b.boldItalic.font);
  // Different requestedFamily still shares faces but has its own wrapper
  const d = await loadFonts(undefined, "Times New Roman");
  assert.notEqual(a, d);
  assert.notEqual(a.metrics, d.metrics);
  assert.equal(
    a.regular.font,
    d.regular.font,
    "faces shared across requestedFamily",
  );
  assert.equal(d.metrics[0].requestedFamily, "Times New Roman");
  assert.equal(d.metrics[0].substitutedMetrics, true);
});

test("bundled loadFonts concurrent calls are single-flight for faces", async () => {
  // Concurrent: Promise.all should coalesce to one getBundledFaces parse for faces,
  // but each wrapper is still fresh (requestedFamily isolation)
  const [p1, p2, p3] = await Promise.all([
    loadFonts(undefined, "Liberation Serif"),
    loadFonts(undefined, "Liberation Serif"),
    loadFonts(undefined, "Liberation Serif"),
  ]);
  assert.notEqual(p1, p2, "concurrent wrappers must be fresh");
  assert.notEqual(p2, p3);
  assert.equal(
    p1.regular.font,
    p2.regular.font,
    "concurrent faces single-flight",
  );
  assert.equal(p2.regular.font, p3.regular.font);
  assert.equal(p1.metrics[0].requestedFamily, "Liberation Serif");
  assert.notEqual(p1.metrics, p2.metrics, "metrics not shared");
  // Across different requestedFamily that share faces
  const [q1, q2] = await Promise.all([
    loadFonts(undefined, "Liberation Serif"),
    loadFonts(undefined, "Times New Roman"),
  ]);
  assert.notEqual(q1, q2);
  assert.equal(
    q1.regular.font,
    q2.regular.font,
    "concurrent across families still single-flight faces",
  );
  assert.equal(q1.metrics[0].requestedFamily, "Liberation Serif");
  assert.equal(q2.metrics[0].requestedFamily, "Times New Roman");
});

test("bundled loadFonts output matches manifest and is stable", async () => {
  const manifest = JSON.parse(
    await readFile(
      new URL(
        "../assets/fonts/liberation-serif-2.1.5/manifest.json",
        import.meta.url,
      ),
      "utf8",
    ),
  );
  const result = await loadFonts(undefined, "Liberation Serif");
  // Output: metricsFamily, sha256, and substitution flag
  assert.equal(result.family, "Liberation Serif");
  assert.deepEqual(
    result.metrics.map(({ role, metricsFamily, sha256 }) => ({
      role,
      metricsFamily,
      sha256,
    })),
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
  assert.equal(
    custom.regular.hash,
    bundled.regular.hash,
    "same bytes → same hash",
  );
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

test("bundled metrics mutation does not contaminate later calls", async () => {
  const first = await loadFonts(undefined, "Liberation Serif");
  const original = first.metrics[0].requestedFamily;
  // Mutate exposed array (simulates caller mutating public measurement result)
  first.metrics.push({});
  first.metrics[0].requestedFamily = "MUTATED";
  first.warnings.push({ code: "MUTATED" });
  const second = await loadFonts(undefined, "Liberation Serif");
  assert.equal(
    second.metrics[0].requestedFamily,
    "Liberation Serif",
    "second call must see original requestedFamily, not mutated",
  );
  assert.equal(
    second.metrics.length,
    4,
    "metrics array must be fresh, not contaminated",
  );
  assert.equal(second.warnings.length, 0, "warnings array must be fresh");
  assert.equal(
    first.metrics[0].requestedFamily,
    "MUTATED",
    "first mutation retained only on first wrapper",
  );
  // Also case-variant must be fresh
  const lower = await loadFonts(undefined, "liberation serif");
  assert.equal(lower.metrics[0].requestedFamily, "liberation serif");
});
