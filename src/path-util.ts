import { isAbsolute, relative, resolve, sep } from "node:path";

/**
 * True when `path` is a normalized relative path: non-empty, NUL-free,
 * backslash-free, not absolute, and composed of no empty/`.`/`..` segments.
 * The single shared predicate behind every unsafe-relative-path check in the
 * codebase; callers keep their own error codes and site-specific extras
 * (prefixes, reserved roots).
 */
export const isSafeRelativePath = (path: string): boolean =>
  path.length > 0 &&
  !path.includes("\0") &&
  !path.includes("\\") &&
  !isAbsolute(path) &&
  path.split("/").every((part) => part !== "" && part !== "." && part !== "..");

/**
 * Canonical cwd-relative display form of a path, with forward separators;
 * "." when the path is the cwd itself.
 */
export const publicPath = (cwd: string, path: string): string => {
  const output = relative(cwd, isAbsolute(path) ? path : resolve(cwd, path))
    .split(sep)
    .join("/");
  return output === "" ? "." : output;
};
