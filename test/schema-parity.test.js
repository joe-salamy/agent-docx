import assert from "node:assert/strict";
import { mkdir, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import test from "node:test";
import { tmpdir } from "node:os";
import { join } from "node:path";
import {
  metadata,
  pkg,
  runInProcess,
  schemaNames,
  validatorFor,
  validatorForSchema,
} from "./helpers.js";
import { compileMarkdown, validateUserRulePack } from "../dist/index.js";
const publishedSchemaNames = [
  "config.schema.json",
  "measurement-request.schema.json",
  "measurement-result.schema.json",
  "docx-template-inspection.schema.json",
  "measurement-stream.schema.json",
  "cli-error.schema.json",
  "profile-catalog.schema.json",
  "project.schema.json",
  "rule-pack.schema.json",
  "agent-request.schema.json",
  "agent-response.schema.json",
  "agent-stream.schema.json",
  "revision.schema.json",
  "change-set.schema.json",
  "source-patch.schema.json",
  "validation-result.schema.json",
  "artifact-result.schema.json",
  "compiled-docx.schema.json",
  "docx-import-result.schema.json",
  "redline-import-result.schema.json",
  "filing-set.schema.json",
  "filing-set-validation.schema.json",
];

const validateProject = validatorFor(
  "https://agent-docx.dev/schemas/project-v1.json",
);
const validateMeasurement = validatorFor(
  "https://agent-docx.dev/schemas/measurement-result-v1.json",
);
const validateValidation = validatorFor(
  "https://agent-docx.dev/schemas/validation-result-v1.json",
);
const validateArtifact = validatorFor(
  "https://agent-docx.dev/schemas/artifact-result-v1.json",
);
const validateDocxImport = validatorFor(
  "https://agent-docx.dev/schemas/docx-import-result-v1.json",
);
const validateRedlineImport = validatorFor(
  "https://agent-docx.dev/schemas/redline-import-result-v1.json",
);
const validateFilingSetValidation = validatorFor(
  "https://agent-docx.dev/schemas/filing-set-validation-v1.json",
);
const validateFilingSet = validatorFor(
  "https://agent-docx.dev/schemas/filing-set-v1.json",
);
const validateRevision = validatorFor(
  "https://agent-docx.dev/schemas/revision-v1.json",
);
const validateChangeSet = validatorFor(
  "https://agent-docx.dev/schemas/change-set-v1.json",
);
const validateRulePack = validatorFor(
  "https://agent-docx.dev/schemas/rule-pack-v1.json",
);
const validateSourcePatch = validatorFor(
  "https://agent-docx.dev/schemas/source-patch-v1.json",
);
const validateCompiledDocx = validatorFor(
  "https://agent-docx.dev/schemas/compiled-docx-v1.json",
);
const validateAgentRequest = validatorFor(
  "https://agent-docx.dev/schemas/agent-request-v1.json",
);
const validateAgentResponse = validatorFor(
  "https://agent-docx.dev/schemas/agent-response-v1.json",
);
const validateFatal = validatorFor(
  "https://agent-docx.dev/schemas/cli-error-v1.json",
);
test("every published schema loads and registers", () => {
  assert.deepEqual(
    [...schemaNames].sort(),
    [...publishedSchemaNames].sort(),
    "published schema set drifted from package exports",
  );
  for (const name of schemaNames)
    assert.equal(
      typeof validatorForSchema(name),
      "function",
      `schema ${name} did not register`,
    );
});

test("happy-path records validate against their published schemas", async () => {
  const directory = await mkdtemp(join(tmpdir(), "agent-docx-parity-"));
  const manifest = "agent-docx.json";
  try {
    await writeFile(
      join(directory, "motion.md"),
      "# Motion\n\nOld statement.\n",
    );
    await writeFile(join(directory, "metadata.json"), JSON.stringify(metadata));

    const measured = await runInProcess(
      ["measure", "--json"],
      "A short filing.\n",
    );
    assert.equal(measured.code, 0, measured.stderr);
    assert.equal(
      validateMeasurement(JSON.parse(measured.stdout)),
      true,
      JSON.stringify(validateMeasurement.errors),
    );

    const initialized = await runInProcess(
      [
        "project",
        "init",
        "--project",
        manifest,
        "--document",
        "motion",
        "--source",
        "motion.md",
        "--profile",
        "us-district-conventional",
        "--metadata",
        "metadata.json",
        "--json",
      ],
      "",
      { cwd: directory },
    );
    assert.equal(initialized.code, 0, initialized.stderr);
    const manifestRecord = JSON.parse(
      await readFile(join(directory, manifest), "utf8"),
    );
    assert.equal(
      validateProject(manifestRecord),
      true,
      JSON.stringify(validateProject.errors),
    );

    const checkpoint = await runInProcess(
      [
        "revision",
        "checkpoint",
        "--project",
        manifest,
        "--document",
        "motion",
        "--author",
        "Drafter",
        "--message",
        "Initial draft",
        "--json",
      ],
      "",
      { cwd: directory },
    );
    assert.equal(checkpoint.code, 0, checkpoint.stderr);
    const checkpointValue = JSON.parse(checkpoint.stdout);
    assert.equal(
      validateRevision(checkpointValue.revision),
      true,
      JSON.stringify(validateRevision.errors),
    );
    const baseRevision = checkpointValue.revision.id;

    await writeFile(
      join(directory, "motion.md"),
      (await readFile(join(directory, "motion.md"), "utf8")).replace(
        "Old statement.",
        "New statement.",
      ),
    );
    const head = await runInProcess(
      [
        "revision",
        "checkpoint",
        "--project",
        manifest,
        "--document",
        "motion",
        "--base",
        baseRevision,
        "--author",
        "Drafter",
        "--message",
        "Revise statement",
        "--json",
      ],
      "",
      { cwd: directory },
    );
    assert.equal(head.code, 0, head.stderr);
    const headRevision = JSON.parse(head.stdout).revision.id;

    const diffed = await runInProcess(
      [
        "revision",
        "diff",
        "--project",
        manifest,
        "--document",
        "motion",
        "--base",
        baseRevision,
        "--head",
        headRevision,
        "--json",
      ],
      "",
      { cwd: directory },
    );
    assert.equal(diffed.code, 0, diffed.stderr);
    assert.equal(
      validateChangeSet(JSON.parse(diffed.stdout).changeSet),
      true,
      JSON.stringify(validateChangeSet.errors),
    );

    const markedSource = await readFile(join(directory, "motion.md"), "utf8");
    const patchStart = markedSource.indexOf("New statement.");
    assert.ok(patchStart >= 0);
    const patch = {
      schemaVersion: 1,
      documentId: "motion",
      baseRevision: headRevision,
      edits: [
        {
          start: patchStart,
          deleteCount: "New statement.".length,
          expectedText: "New statement.",
          replacement: "Revised statement.",
        },
      ],
    };
    await writeFile(join(directory, "patch.json"), JSON.stringify(patch));
    const evaluated = await runInProcess(
      [
        "draft",
        "evaluate",
        "--project",
        manifest,
        "--document",
        "motion",
        "--patch",
        "patch.json",
        "--json",
      ],
      "",
      { cwd: directory },
    );
    assert.equal(evaluated.code, 0, evaluated.stderr);
    assert.equal(
      validateSourcePatch(patch),
      true,
      JSON.stringify(validateSourcePatch.errors),
    );

    const validation = await runInProcess(
      ["validate", "--project", manifest, "--document", "motion", "--json"],
      "",
      { cwd: directory },
    );
    assert.equal(validation.code, 0, validation.stderr);
    assert.equal(
      validateValidation(JSON.parse(validation.stdout)),
      true,
      JSON.stringify(validateValidation.errors),
    );

    const exported = await runInProcess(
      [
        "export",
        "--project",
        manifest,
        "--document",
        "motion",
        "--revision",
        "HEAD",
        "--mode",
        "clean",
        "--output",
        "motion.docx",
        "--json",
      ],
      "",
      { cwd: directory },
    );
    assert.equal(exported.code, 0, exported.stderr);
    const exportValue = JSON.parse(exported.stdout);
    assert.equal(
      validateArtifact(exportValue.artifact),
      true,
      JSON.stringify(validateArtifact.errors),
    );

    const redline = await runInProcess(
      [
        "export",
        "--project",
        manifest,
        "--document",
        "motion",
        "--revision",
        "HEAD",
        "--mode",
        "redline",
        "--base",
        baseRevision,
        "--output",
        "motion-redline.docx",
        "--json",
      ],
      "",
      { cwd: directory },
    );
    assert.equal(redline.code, 0, redline.stderr);

    const imported = await runInProcess(
      ["import", "motion.docx", "--inspect-only", "--json"],
      "",
      { cwd: directory },
    );
    assert.equal(imported.code, 0, imported.stderr);
    assert.equal(
      validateDocxImport(JSON.parse(imported.stdout)),
      true,
      JSON.stringify(validateDocxImport.errors),
    );

    const redlineImported = await runInProcess(
      [
        "import-redline",
        "--project",
        manifest,
        "--document",
        "motion",
        "--input",
        "motion-redline.docx",
        "--author",
        "Drafter",
        "--message",
        "Redline import round trip",
        "--json",
      ],
      "",
      { cwd: directory },
    );
    assert.equal(redlineImported.code, 0, redlineImported.stderr);
    assert.equal(
      validateRedlineImport(JSON.parse(redlineImported.stdout)),
      true,
      JSON.stringify(validateRedlineImport.errors),
    );

    const filingSet = await runInProcess(
      [
        "filing-set",
        "add",
        "--project",
        manifest,
        "--id",
        "set-1",
        "--documents",
        "motion",
        "--json",
      ],
      "",
      { cwd: directory },
    );
    assert.equal(filingSet.code, 0, filingSet.stderr);
    const filingSetValue = JSON.parse(filingSet.stdout);
    assert.equal(filingSetValue.filingSets.length, 1);

    const filingSetGot = await runInProcess(
      ["filing-set", "get", "--project", manifest, "--id", "set-1", "--json"],
      "",
      { cwd: directory },
    );
    assert.equal(filingSetGot.code, 0, filingSetGot.stderr);
    assert.equal(
      validateFilingSet(JSON.parse(filingSetGot.stdout)),
      true,
      JSON.stringify(validateFilingSet.errors),
    );

    const filingSetValidated = await runInProcess(
      [
        "filing-set",
        "validate",
        "--project",
        manifest,
        "--id",
        "set-1",
        "--json",
      ],
      "",
      { cwd: directory },
    );
    assert.equal(filingSetValidated.code, 0, filingSetValidated.stderr);
    assert.equal(
      validateFilingSetValidation(JSON.parse(filingSetValidated.stdout)),
      true,
      JSON.stringify(validateFilingSetValidation.errors),
    );

    const agentRequests = [
      {
        schemaVersion: 1,
        id: "parity-1",
        action: "project.get",
        params: {},
      },
      {
        schemaVersion: 1,
        id: "parity-2",
        action: "document.measure",
        project: "agent-docx.json",
        params: { documentId: "motion" },
      },
      {
        schemaVersion: 1,
        id: "parity-3",
        action: "document.validate",
        project: manifest,
        params: { documentId: "motion" },
      },
    ];
    for (const request of agentRequests)
      assert.equal(
        validateAgentRequest(request),
        true,
        JSON.stringify(validateAgentRequest.errors),
      );
    const agentSession = await runInProcess(
      ["agent", "--input-jsonl"],
      agentRequests.map((request) => JSON.stringify(request)).join("\n"),
      { cwd: directory },
    );
    assert.equal(agentSession.code, 0, agentSession.stderr);
    const agentRecords = agentSession.stdout
      .trim()
      .split("\n")
      .filter((line) => line.length > 0)
      .map((line) => JSON.parse(line));
    assert.equal(agentRecords.length, 3);
    for (const record of agentRecords) {
      assert.equal(record.kind, "result");
      assert.equal(
        validateAgentResponse(record),
        true,
        JSON.stringify(validateAgentResponse.errors),
      );
    }
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});

test("fatal records validate against the CLI error schema", async () => {
  const directory = await mkdtemp(join(tmpdir(), "agent-docx-fatal-"));
  const manifest = "agent-docx.json";
  try {
    const missing = await runInProcess(["validate"], "", { cwd: directory });
    assert.equal(missing.code, 1);
    const missingRecord = JSON.parse(missing.stderr);
    assert.equal(missingRecord.error.code, "PROJECT_NOT_FOUND");
    assert.equal(
      validateFatal(missingRecord),
      true,
      JSON.stringify(validateFatal.errors),
    );

    const badFlag = await runInProcess(["measure", "--bogus"], "", {
      cwd: directory,
    });
    assert.equal(badFlag.code, 2);
    const flagRecord = JSON.parse(badFlag.stderr);
    assert.equal(flagRecord.error.code, "INVALID_ARGUMENT");
    assert.equal(
      validateFatal(flagRecord),
      true,
      JSON.stringify(validateFatal.errors),
    );

    await writeFile(join(directory, "motion.md"), "# Motion\n\nBody.\n");
    await writeFile(join(directory, "metadata.json"), JSON.stringify(metadata));
    const initialized = await runInProcess(
      [
        "project",
        "init",
        "--project",
        manifest,
        "--document",
        "motion",
        "--source",
        "motion.md",
        "--profile",
        "us-district-conventional",
        "--metadata",
        "metadata.json",
      ],
      "",
      { cwd: directory },
    );
    assert.equal(initialized.code, 0, initialized.stderr);

    await mkdir(join(directory, "packs"), { recursive: true });
    await writeFile(join(directory, "packs", "bad.json"), "{not json");
    await writeFile(
      join(directory, "changes.json"),
      JSON.stringify({ rulePacks: ["packs/bad.json"] }),
    );
    const configure = await runInProcess(
      [
        "document",
        "configure",
        "--project",
        manifest,
        "--document",
        "motion",
        "--base",
        "HEAD",
        "--changes",
        "changes.json",
        "--author",
        "Drafter",
        "--message",
        "Bind a broken pack",
      ],
      "",
      { cwd: directory },
    );
    assert.equal(configure.code, 1);
    const packRecord = JSON.parse(configure.stderr);
    assert.equal(packRecord.error.code, "RULE_PACK_INVALID");
    assert.equal(
      validateFatal(packRecord),
      true,
      JSON.stringify(validateFatal.errors),
    );

    const first = await runInProcess(
      [
        "revision",
        "checkpoint",
        "--project",
        manifest,
        "--document",
        "motion",
        "--author",
        "Drafter",
        "--message",
        "First",
        "--json",
      ],
      "",
      { cwd: directory },
    );
    assert.equal(first.code, 0, first.stderr);
    const firstRevision = JSON.parse(first.stdout).revision.id;
    const second = await runInProcess(
      [
        "revision",
        "checkpoint",
        "--project",
        manifest,
        "--document",
        "motion",
        "--base",
        firstRevision,
        "--author",
        "Drafter",
        "--message",
        "Second",
        "--json",
      ],
      "",
      { cwd: directory },
    );
    assert.equal(second.code, 0, second.stderr);
    const stale = await runInProcess(
      [
        "revision",
        "checkpoint",
        "--project",
        manifest,
        "--document",
        "motion",
        "--base",
        firstRevision,
        "--author",
        "Drafter",
        "--message",
        "Stale base",
        "--json",
      ],
      "",
      { cwd: directory },
    );
    assert.equal(stale.code, 1);
    const staleRecord = JSON.parse(stale.stderr);
    assert.equal(staleRecord.error.code, "REVISION_CONFLICT");
    assert.equal(
      validateFatal(staleRecord),
      true,
      JSON.stringify(validateFatal.errors),
    );
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});

test("serializable compile results validate against compiled-docx schema", async () => {
  const compiled = await compileMarkdown("# Motion\n\nRequested relief.\n", {
    documentId: "motion",
    profile: "us-district-conventional",
    metadata,
  });
  const serializable = {
    schemaVersion: compiled.schemaVersion,
    measurement: compiled.measurement,
    validation: compiled.validation,
    blocks: compiled.blocks,
    artifact: compiled.artifact,
  };
  assert.equal(
    validateCompiledDocx(serializable),
    true,
    JSON.stringify(validateCompiledDocx.errors),
  );
});

test("manifest and rule-pack round trips match runtime acceptance", async () => {
  const manifest = {
    schemaVersion: 1,
    projectId: "11111111-2222-4333-8444-555555555555",
    defaultDocument: "motion",
    storeDir: ".agent-docx",
    documents: [
      {
        id: "motion",
        source: "motion.md",
        profile: "us-district-conventional",
        metadata,
        rulePacks: ["rules/custom.json"],
      },
    ],
  };
  assert.equal(
    validateProject(manifest),
    true,
    JSON.stringify(validateProject.errors),
  );

  const pack = {
    id: "custom-rule@2026-08-01",
    sourceUrl: "https://example.com/rule.txt",
    effectiveDate: "2026-08-01",
    sourceSha256: `sha256:${"a".repeat(64)}`,
    sourceExcerpt: "Rule text excerpt",
    checks: [
      {
        id: "custom.lines",
        kind: "counted-lines-maximum",
        citation: "Local Rule 1",
        predicate: "No more than 25 counted lines per page",
        params: { perPageMaximum: 25 },
      },
    ],
    unmodeledProvisions: [],
  };
  assert.equal(
    validateRulePack(pack),
    true,
    JSON.stringify(validateRulePack.errors),
  );
  const accepted = validateUserRulePack(pack);
  assert.equal(accepted.checks[0].params.perPageMaximum, 25);
});
test("schema contracts cover empty results, custom packs, and rejected malformed records", async () => {
  const measured = await runInProcess(
    ["measure", "--json"],
    "A short filing.\n",
  );
  assert.equal(measured.code, 0, measured.stderr);
  const emptyMeasurement = JSON.parse(measured.stdout);
  emptyMeasurement.pageCount = 0;
  emptyMeasurement.deterministic.pageCount = 0;
  emptyMeasurement.deterministic.equivalentPages = 0;
  emptyMeasurement.deterministic.totalVisualLines = 0;
  emptyMeasurement.deterministic.visualLinesByPage = [];
  emptyMeasurement.deterministic.countedLinesByPage = [];
  emptyMeasurement.deterministic.lastPage = null;
  assert.equal(
    validateMeasurement(emptyMeasurement),
    true,
    JSON.stringify(validateMeasurement.errors),
  );

  const customValidation = {
    schemaVersion: 1,
    documentId: "motion",
    revision: null,
    rulePack: null,
    scope: {
      certification: false,
      checkedRuleIds: [],
      sourceSnapshots: [
        {
          id: "local-rules@2026-01-01",
          sourceUrl: "https://example.com/local-rules.txt",
          effectiveDate: "2026-01-01",
          sha256: `sha256:${"b".repeat(64)}`,
        },
      ],
      unmodeledProvisions: [],
    },
    status: "unknown",
    summary: { pass: 0, fail: 0, unknown: 0 },
    findings: [],
  };
  assert.equal(
    validateValidation(customValidation),
    true,
    JSON.stringify(validateValidation.errors),
  );

  const revision = `sha256:${"c".repeat(64)}`;
  // The patch contract expresses deletion length as a nonnegative
  // deleteCount, so an inverted range is structurally inexpressible; the
  // schema itself rejects negative counts.
  assert.equal(
    validateSourcePatch({
      schemaVersion: 1,
      documentId: "motion",
      baseRevision: revision,
      edits: [{ start: 4, deleteCount: -1, expectedText: "", replacement: "" }],
    }),
    false,
    "negative deleteCount must be rejected by the schema",
  );
  assert.equal(
    validateSourcePatch({
      schemaVersion: 1,
      documentId: "motion",
      baseRevision: revision,
      edits: [{ start: 4, deleteCount: 3, expectedText: "", replacement: "" }],
    }),
    true,
    "nonnegative deleteCount must be accepted by the schema",
  );

  const agentProjectParams = {
    documentId: "motion",
    source: "motion.md",
    profile: "us-district-conventional",
    metadata,
    makeDefault: true,
  };
  assert.equal(
    validateAgentRequest({
      schemaVersion: 1,
      action: "project.add",
      params: agentProjectParams,
    }),
    true,
    JSON.stringify(validateAgentRequest.errors),
  );
  assert.equal(
    validateAgentRequest({
      schemaVersion: 1,
      action: "project.init",
      params: agentProjectParams,
    }),
    false,
    "project.init must reject makeDefault",
  );

  const invalidAnnotationChange = {
    schemaVersion: 1,
    id: revision,
    documentId: "motion",
    baseRevision: revision,
    headRevision: revision,
    changes: [],
    annotations: [
      {
        id: `c_${"d".repeat(64)}`,
        kind: "add",
        newValue: {
          id: "a_00000000-0000-0000-0000-000000000000",
          blockId: "b_00000000-0000-4000-8000-000000000000",
          author: null,
          createdAt: null,
          message: "Invalid UUID version",
          status: "open",
        },
      },
    ],
  };
  assert.equal(
    validateChangeSet(invalidAnnotationChange),
    false,
    "annotation IDs must be UUID v4 with the RFC variant",
  );

  const artifactWithoutPdf = {
    schemaVersion: 1,
    mediaType:
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    byteLength: 1,
    sha256: revision,
    provenanceSha256: revision,
    documentId: "motion",
    profile: "us-district-conventional",
    rulePack: null,
    rendererProvenance: {
      generator: "agent-docx",
      requested: "deterministic",
      pageCountSource: "deterministic",
    },
    path: "motion.docx",
    storePath: ".agent-docx/objects/docx",
    attachments: null,
    revision,
    mode: "pdf",
    baseRevision: null,
  };
  assert.equal(
    validateArtifact(artifactWithoutPdf),
    false,
    "PDF artifacts must include their PDF payload",
  );

  const duplicateManifest = {
    schemaVersion: 1,
    projectId: "11111111-2222-4333-8444-555555555555",
    defaultDocument: "motion",
    storeDir: ".agent-docx",
    documents: [
      {
        id: "motion",
        source: "motion.md",
        profile: "us-district-conventional",
        metadata,
      },
      {
        id: "motion",
        source: "motion.md",
        profile: "us-district-conventional",
        metadata,
      },
    ],
  };
  assert.equal(
    validateProject(duplicateManifest),
    false,
    "duplicate document IDs must be rejected",
  );
});

test("published package version matches the runtime version literal", () => {
  assert.equal(pkg.version, "0.1.0");
});
