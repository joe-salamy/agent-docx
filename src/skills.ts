import { cp, mkdir, readdir, readFile, rm, stat } from "node:fs/promises";
import { dirname, isAbsolute, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { homedir } from "node:os";
import { AgentDocxError } from "./types.js";

export type SkillInfo = {
  name: string;
  description: string;
  version: string;
  sourcePath: string;
};

export type InstallResult = {
  name: string;
  sourcePath: string;
  destPath: string;
  status: "installed" | "overwritten" | "skipped" | "dry-run";
};

async function resolveSourceDir(): Promise<string> {
  // dist/skills.js -> ../skills, src/skills.ts -> ../skills
  const here = dirname(fileURLToPath(import.meta.url));
  const candidates = [
    resolve(here, "../skills"),
    resolve(here, "../../skills"),
    resolve(here, "skills"),
  ];
  for (const candidate of candidates) {
    try {
      const st = await stat(candidate);
      if (st.isDirectory()) return candidate;
    } catch {
      // try next
    }
  }
  throw new AgentDocxError(
    "INTERNAL_ERROR",
    `Skills source directory not found (tried ${candidates.join(", ")})`,
  );
}

function parseFrontmatter(
  text: string,
): { name?: string; description?: string; version?: string } {
  const match = text.match(/^---\r?\n([\s\S]*?)\r?\n---/);
  if (!match) return {};
  const yaml = match[1] ?? "";
  const out: Record<string, string> = {};
  for (const line of yaml.split(/\r?\n/)) {
    const idx = line.indexOf(":");
    if (idx === -1) continue;
    const key = line.slice(0, idx).trim();
    let value = line.slice(idx + 1).trim();
    // strip quotes
    if (
      (value.startsWith('"') && value.endsWith('"')) ||
      (value.startsWith("'") && value.endsWith("'"))
    ) {
      value = value.slice(1, -1);
    }
    out[key] = value;
  }
  return out;
}

export async function listSkills(): Promise<SkillInfo[]> {
  const sourceDir = await resolveSourceDir();
  const entries = await readdir(sourceDir, { withFileTypes: true });
  const skills: SkillInfo[] = [];
  for (const entry of entries) {
    if (!entry.isDirectory()) continue;
    const skillDir = join(sourceDir, entry.name);
    const skillFile = join(skillDir, "SKILL.md");
    try {
      const st = await stat(skillFile);
      if (!st.isFile()) continue;
    } catch {
      continue;
    }
    const text = await readFile(skillFile, "utf8");
    const fm = parseFrontmatter(text);
    skills.push({
      name: fm.name ?? entry.name,
      description: fm.description ?? "",
      version: fm.version ?? "0.0.0",
      sourcePath: skillDir,
    });
  }
  skills.sort((a, b) => a.name.localeCompare(b.name));
  return skills;
}

export async function installSkills(options: {
  cwd: string;
  dest?: string;
  global?: boolean;
  force?: boolean;
  dryRun?: boolean;
}): Promise<{
  destBase: string;
  results: InstallResult[];
  sourceDir: string;
}> {
  if (options.dest && options.global) {
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      "--dest and --global cannot be combined",
    );
  }
  const sourceDir = await resolveSourceDir();
  const skills = await listSkills();
  if (skills.length === 0) {
    throw new AgentDocxError("INTERNAL_ERROR", "No skills found in package");
  }
  let destBase: string;
  if (options.global) {
    destBase = join(homedir(), ".omp", "skills");
  } else if (options.dest) {
    destBase = isAbsolute(options.dest)
      ? options.dest
      : resolve(options.cwd, options.dest);
  } else {
    destBase = resolve(options.cwd, ".omp", "skills");
  }

  const results: InstallResult[] = [];
  if (!options.dryRun) {
    await mkdir(destBase, { recursive: true });
  }
  for (const skill of skills) {
    const destPath = join(destBase, skill.name);
    let exists = false;
    try {
      const st = await stat(destPath);
      exists = !!st;
    } catch {
      exists = false;
    }
    if (exists && !options.force && !options.dryRun) {
      // report skipped, require --force
      results.push({
        name: skill.name,
        sourcePath: skill.sourcePath,
        destPath,
        status: "skipped",
      });
      continue;
    }
    if (options.dryRun) {
      results.push({
        name: skill.name,
        sourcePath: skill.sourcePath,
        destPath,
        status: "dry-run",
      });
      continue;
    }
    if (exists && options.force) {
      await rm(destPath, { recursive: true, force: true });
    }
    await cp(skill.sourcePath, destPath, { recursive: true, force: true });
    results.push({
      name: skill.name,
      sourcePath: skill.sourcePath,
      destPath,
      status: exists && options.force ? "overwritten" : "installed",
    });
  }

  // If any skipped due to existing without --force, surface an error after partial?
  // We allow partial and let caller report; but if you want strict, throw.
  const skipped = results.filter((r) => r.status === "skipped");
  if (skipped.length > 0) {
    // Do not throw by default; let caller inspect. We'll throw only if all skipped? Instead we attach info.
    // Throw with details so CLI can print friendly message and set exit code 1 but still show results.
    // For now, we don't throw — CLI layer will decide.
  }

  return { destBase, results, sourceDir };
}
