import { glob, lstat, realpath, stat } from "node:fs/promises";
import type { Stats } from "node:fs";
import { isAbsolute, relative, resolve, sep } from "node:path";
import { AgentDocxError } from "../types.js";
import type { BatchSelection, Source } from "../cli-contract.js";

function invalidPattern(pattern: string) {
  if (!pattern || isAbsolute(pattern) || pattern.startsWith("!")) return true;
  let brackets = 0;
  let escaped = false;
  for (const character of pattern) {
    if (escaped) {
      escaped = false;
    } else if (character === "\\") {
      escaped = true;
    } else if (character === "[") {
      brackets++;
    } else if (character === "]") {
      if (brackets === 0) return true;
      brackets--;
    }
  }
  return brackets !== 0;
}

export const normalizedRelativePath = (cwd: string, path: string) =>
  relative(cwd, path).split(sep).join("/");

export async function resolveBatchInputs(
  selectors: readonly string[],
  selection: BatchSelection,
  cwd: string,
): Promise<Extract<Source, { kind: "file" }>[]> {
  for (const pattern of [...selection.include, ...selection.exclude]) {
    if (invalidPattern(pattern)) {
      throw new AgentDocxError(
        "INVALID_ARGUMENT",
        `Invalid batch pattern: ${pattern}`,
      );
    }
  }
  const output: Extract<Source, { kind: "file" }>[] = [];
  const seen = new Set<string>();
  const add = async (
    source: Extract<Source, { kind: "file" }>,
    existing: boolean,
  ) => {
    const identity = existing
      ? await realpath(source.resolvedPath)
      : source.resolvedPath;
    if (seen.has(identity)) return;
    seen.add(identity);
    output.push(source);
  };
  const discoveredSources = async (
    pattern: string | readonly string[],
    root: string,
    excludes: readonly string[],
    errorPattern: string,
    directOnly: boolean,
  ) => {
    const candidates: Extract<Source, { kind: "file" }>[] = [];
    try {
      for await (const matched of glob(pattern, {
        cwd: root,
        exclude: excludes,
      })) {
        const resolvedPath = resolve(root, matched);
        const sourcePath = normalizedRelativePath(cwd, resolvedPath);
        const matchedPath = String(matched).split(sep).join("/");
        if (directOnly && matchedPath.includes("/")) continue;
        try {
          if (!(await stat(resolvedPath)).isFile()) continue;
        } catch {
          continue;
        }
        candidates.push({
          kind: "file",
          path: sourcePath,
          resolvedPath,
        });
      }
    } catch {
      throw new AgentDocxError(
        "INVALID_ARGUMENT",
        `Invalid batch pattern: ${errorPattern}`,
      );
    }
    candidates.sort((a, b) => (a.path < b.path ? -1 : a.path > b.path ? 1 : 0));
    const unique: Extract<Source, { kind: "file" }>[] = [];
    const localSeen = new Set<string>();
    for (const candidate of candidates) {
      const identity = await realpath(candidate.resolvedPath);
      if (localSeen.has(identity)) continue;
      localSeen.add(identity);
      unique.push(candidate);
    }
    return unique;
  };

  for (const selector of selectors) {
    if (!selector) {
      throw new AgentDocxError(
        "INVALID_ARGUMENT",
        "Batch selector must not be empty",
      );
    }
    if (selector === "-") {
      throw new AgentDocxError(
        "INVALID_ARGUMENT",
        "Positional batch does not accept stdin; use --input-jsonl",
      );
    }
    const resolvedSelector = resolve(cwd, selector);
    let information: Stats | undefined;
    try {
      information = await lstat(resolvedSelector);
    } catch (error) {
      if (
        !(
          error &&
          typeof error === "object" &&
          "code" in error &&
          error.code === "ENOENT"
        )
      ) {
        throw error;
      }
    }
    if (information) {
      if (information.isSymbolicLink()) {
        try {
          if ((await stat(resolvedSelector)).isFile()) {
            await add(
              {
                kind: "file",
                path: selector,
                resolvedPath: resolvedSelector,
              },
              true,
            );
            continue;
          }
        } catch {}
        throw new AgentDocxError(
          "INVALID_ARGUMENT",
          `Unsupported batch input: ${selector}`,
        );
      }
      if (information.isFile()) {
        await add(
          { kind: "file", path: selector, resolvedPath: resolvedSelector },
          true,
        );
        continue;
      }
      if (information.isDirectory()) {
        const include = selection.recursive
          ? selection.include.map((pattern) =>
              pattern.includes("/") ? pattern : `**/${pattern}`,
            )
          : selection.include.filter((pattern) => !pattern.includes("/"));
        const exclude = selection.exclude.map((pattern) =>
          selection.recursive && !pattern.includes("/")
            ? `**/${pattern}`
            : pattern,
        );
        const matches =
          include.length === 0
            ? []
            : await discoveredSources(
                include,
                resolvedSelector,
                exclude,
                include[0]!,
                !selection.recursive,
              );
        if (matches.length === 0) {
          throw new AgentDocxError(
            "INVALID_ARGUMENT",
            `Batch selector matched no files: ${selector}`,
          );
        }
        for (const source of matches) await add(source, true);
        continue;
      }
      throw new AgentDocxError(
        "INVALID_ARGUMENT",
        `Unsupported batch input: ${selector}`,
      );
    }

    if (selector.startsWith("!")) {
      throw new AgentDocxError(
        "INVALID_ARGUMENT",
        `Invalid batch pattern: ${selector}`,
      );
    }
    if (/[*?[\]]/.test(selector)) {
      if (invalidPattern(selector)) {
        throw new AgentDocxError(
          "INVALID_ARGUMENT",
          `Invalid batch pattern: ${selector}`,
        );
      }
      const exclude = selection.exclude.map((pattern) =>
        pattern.includes("/") ? pattern : `**/${pattern}`,
      );
      const matches = await discoveredSources(
        selector,
        cwd,
        exclude,
        selector,
        false,
      );
      if (matches.length === 0) {
        throw new AgentDocxError(
          "INVALID_ARGUMENT",
          `Batch selector matched no files: ${selector}`,
        );
      }
      for (const source of matches) await add(source, true);
      continue;
    }

    await add(
      { kind: "file", path: selector, resolvedPath: resolvedSelector },
      false,
    );
  }
  return output;
}
