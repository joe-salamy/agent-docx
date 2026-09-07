import type { ChangeSet } from "../revisions/types.js";
import type { SourcePatch } from "../draft/types.js";
import type {
  AgentAction,
  AgentParamsForAction,
  AgentRequest,
  AgentRequestId,
} from "./protocol-primitives.js";
import { AgentDocxError } from "../types.js";
import {
  assertConfigUpdate,
  assertProjectInput,
  assertRendererOptions,
} from "./protocol-metadata.js";
import {
  actor,
  agentActions,
  annotationIdPattern,
  asRevision,
  assertKeys,
  assertRecord,
  blockIdPattern,
  changeIdPattern,
  documentIdPattern,
  hasOwn,
  noOptions,
  optionalEnum,
  optionalInteger,
  optionalRevision,
  optionalString,
  requiredDocumentId,
  requiredEnum,
  requiredInteger,
  requiredRevisionId,
  requiredString,
  requiredText,
  rendererModes,
} from "./protocol-primitives.js";

export const asPatch = (value: unknown): SourcePatch => {
  assertSourcePatch(value);
  return value as SourcePatch;
};

export const asChangeSet = (value: unknown): ChangeSet => {
  assertChangeSet(value);
  return value as ChangeSet;
};

export const assertSourcePatch = (value: unknown): void => {
  const patch = assertRecord(value, "patch");
  assertKeys(
    patch,
    ["schemaVersion", "documentId", "baseRevision", "edits"],
    "patch",
  );
  if (patch.schemaVersion !== 1)
    throw new AgentDocxError("PATCH_INVALID", "patch.schemaVersion must be 1");
  requiredDocumentId(patch, "documentId");
  requiredRevisionId(patch, "baseRevision");
  if (!Array.isArray(patch.edits))
    throw new AgentDocxError("PATCH_INVALID", "patch.edits must be an array");
  for (const edit of patch.edits) {
    const record = assertRecord(edit, "patch edit");
    assertKeys(
      record,
      ["start", "deleteCount", "expectedText", "replacement"],
      "patch edit",
    );
    requiredInteger(record, "start");
    requiredInteger(record, "deleteCount");
    if ((record.deleteCount as number) < 0)
      throw new AgentDocxError(
        "PATCH_INVALID",
        "patch edit deleteCount must not be negative",
      );
    requiredText(record, "expectedText", "patch edit");
    requiredText(record, "replacement", "patch edit");
  }
};

export const assertChangeSet = (value: unknown): void => {
  const changeSet = assertRecord(value, "changeSet");
  assertKeys(
    changeSet,
    [
      "schemaVersion",
      "id",
      "documentId",
      "baseRevision",
      "headRevision",
      "changes",
      "annotations",
    ],
    "changeSet",
  );
  if (changeSet.schemaVersion !== 1)
    throw new AgentDocxError(
      "CHANGESET_INVALID",
      "changeSet.schemaVersion must be 1",
    );
  requiredRevisionId(changeSet, "id");
  requiredDocumentId(changeSet, "documentId");
  requiredRevisionId(changeSet, "baseRevision");
  requiredRevisionId(changeSet, "headRevision");
  if (
    !Array.isArray(changeSet.changes) ||
    !Array.isArray(changeSet.annotations)
  )
    throw new AgentDocxError(
      "CHANGESET_INVALID",
      "changeSet changes and annotations must be arrays",
    );
  for (const [index, change] of changeSet.changes.entries())
    assertChangeItem(change, `changeSet.changes[${index}]`);
  for (const [index, annotation] of changeSet.annotations.entries())
    assertAnnotationChange(annotation, `changeSet.annotations[${index}]`);
};

const changeKinds = new Set([
  "insert-block",
  "delete-block",
  "move-block",
  "replace-block",
  "replace-container-shell",
  "insert-text",
  "delete-text",
  "replace-text",
  "add-config",
  "remove-config",
  "replace-config",
  "add-dependency",
  "remove-dependency",
  "replace-dependency",
]);

const dateTimePattern =
  /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})(?:\.\d+)?(Z|[+-](\d{2}):(\d{2}))$/;

const assertDateTime = (value: string, label: string): void => {
  const match = dateTimePattern.exec(value);
  const year = match ? Number(match[1]) : NaN;
  const month = match ? Number(match[2]) : NaN;
  const day = match ? Number(match[3]) : NaN;
  const hour = match ? Number(match[4]) : NaN;
  const minute = match ? Number(match[5]) : NaN;
  const second = match ? Number(match[6]) : NaN;
  const offsetHour = match?.[8] === undefined ? 0 : Number(match[8]);
  const offsetMinute = match?.[9] === undefined ? 0 : Number(match[9]);
  const leapYear = year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0);
  const daysInMonth =
    month === 2
      ? leapYear
        ? 29
        : 28
      : [4, 6, 9, 11].includes(month)
        ? 30
        : 31;
  if (
    !match ||
    month < 1 ||
    month > 12 ||
    day < 1 ||
    day > daysInMonth ||
    hour > 23 ||
    minute > 59 ||
    second > 59 ||
    offsetHour > 23 ||
    offsetMinute > 59
  )
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `${label} must be an RFC 3339 date-time or null`,
    );
};
const assertChangeSetActor = (value: unknown, label: string): void => {
  const author = assertRecord(value, label);
  assertKeys(author, ["name", "email"], label);
  if (typeof author.name !== "string" || author.name.length === 0)
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `${label}.name must be a non-empty string`,
    );
  if (author.email !== undefined && typeof author.email !== "string")
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `${label}.email must be a string`,
    );
};

export const assertAttribution = (value: unknown, label: string): void => {
  const attribution = assertRecord(value, `${label}.attribution`);
  assertKeys(
    attribution,
    ["author", "createdAt", "sourceRevisionId"],
    `${label}.attribution`,
  );
  if (!("author" in attribution) || !("createdAt" in attribution))
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `${label}.attribution must have author and createdAt`,
    );
  if (attribution.author !== null)
    assertChangeSetActor(attribution.author, `${label}.attribution.author`);
  if (attribution.createdAt !== null) {
    if (typeof attribution.createdAt !== "string")
      throw new AgentDocxError(
        "INVALID_ARGUMENT",
        `${label}.attribution.createdAt must be a date-time or null`,
      );
    assertDateTime(attribution.createdAt, `${label}.attribution.createdAt`);
  }
  if (
    attribution.sourceRevisionId !== undefined &&
    typeof attribution.sourceRevisionId !== "string"
  )
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `${label}.attribution.sourceRevisionId must be a string`,
    );
};

export const assertBlockLocation = (value: unknown, label: string): void => {
  const location = assertRecord(value, label);
  assertKeys(
    location,
    ["collection", "parentId", "index", "sourceOffset"],
    label,
  );
  if (location.collection !== "body" && location.collection !== "footnotes")
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `${label}.collection must be body or footnotes`,
    );
  if (
    location.parentId !== null &&
    (typeof location.parentId !== "string" ||
      !blockIdPattern.test(location.parentId))
  )
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `${label}.parentId must be a block ID or null`,
    );
  requiredInteger(location, "index");
  requiredInteger(location, "sourceOffset");
};

export const assertSourceRange = (value: unknown, label: string): void => {
  const range = assertRecord(value, label);
  assertKeys(range, ["start", "end", "text"], label);
  requiredInteger(range, "start");
  requiredInteger(range, "end");
  if ((range.end as number) < (range.start as number))
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `${label}.end must not precede start`,
    );
  requiredText(range, "text", label);
};
const assertAttributionSpans = (value: unknown, label: string): void => {
  if (!Array.isArray(value))
    throw new AgentDocxError("INVALID_ARGUMENT", `${label} must be an array`);
  for (const [index, entry] of value.entries()) {
    const span = assertRecord(entry, `${label}[${index}]`);
    assertKeys(span, ["start", "end", "attribution"], `${label}[${index}]`);
    const start = requiredInteger(span, "start");
    const end = requiredInteger(span, "end");
    if (end < start)
      throw new AgentDocxError(
        "INVALID_ARGUMENT",
        `${label}[${index}].end must not precede start`,
      );
    assertAttribution(span.attribution, `${label}[${index}]`);
  }
};

export const assertChangeItem = (value: unknown, label: string): void => {
  const change = assertRecord(value, label);
  if (
    Object.keys(change).some(
      (key) =>
        ![
          "id",
          "kind",
          "attribution",
          "blockId",
          "from",
          "to",
          "oldSource",
          "newSource",
          "oldOffset",
          "newOffset",
          "oldText",
          "newText",
          "oldAttributionSpans",
          "newAttributionSpans",
          "block",
          "oldBlock",
          "newBlock",
          "oldShell",
          "newShell",
          "path",
          "oldValue",
          "newValue",
          "key",
          "oldObject",
          "newObject",
        ].includes(key),
    )
  )
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `${label} has an unsupported property`,
    );
  if (typeof change.id !== "string" || !changeIdPattern.test(change.id))
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `${label}.id must be a canonical change ID`,
    );
  if (typeof change.kind !== "string" || !changeKinds.has(change.kind))
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `${label}.kind is not a supported change kind`,
    );
  assertAttribution(change.attribution, label);
  if (change.blockId !== undefined) {
    if (
      typeof change.blockId !== "string" ||
      !blockIdPattern.test(change.blockId)
    )
      throw new AgentDocxError(
        "INVALID_ARGUMENT",
        `${label}.blockId must be a block ID`,
      );
  }
  if (change.from !== undefined)
    assertBlockLocation(change.from, `${label}.from`);
  if (change.to !== undefined) assertBlockLocation(change.to, `${label}.to`);
  if (change.oldSource !== undefined)
    assertSourceRange(change.oldSource, `${label}.oldSource`);
  if (change.newSource !== undefined)
    assertSourceRange(change.newSource, `${label}.newSource`);
  for (const key of ["oldAttributionSpans", "newAttributionSpans"] as const)
    if (change[key] !== undefined)
      assertAttributionSpans(change[key], `${label}.${key}`);
  for (const key of ["oldOffset", "newOffset"] as const)
    if (change[key] !== undefined) requiredInteger(change, key);
};

export const assertAnnotationChange = (value: unknown, label: string): void => {
  const change = assertRecord(value, label);
  assertKeys(change, ["id", "kind", "oldValue", "newValue"], label);
  if (typeof change.id !== "string" || !changeIdPattern.test(change.id))
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `${label}.id must be a canonical change ID`,
    );
  if (
    change.kind !== "add" &&
    change.kind !== "replace" &&
    change.kind !== "remove"
  )
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `${label}.kind must be add, replace, or remove`,
    );
  if (change.kind === "add" && change.newValue === undefined)
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `${label} must carry newValue`,
    );
  if (change.kind === "remove" && change.oldValue === undefined)
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `${label} must carry oldValue`,
    );
  for (const key of ["oldValue", "newValue"] as const)
    if (change[key] !== undefined)
      assertReviewAnnotation(change[key]!, `${label}.${key}`);
};

export const assertReviewAnnotation = (value: unknown, label: string): void => {
  const annotation = assertRecord(value, label);
  assertKeys(
    annotation,
    ["id", "blockId", "range", "author", "createdAt", "message", "status"],
    label,
  );
  for (const key of [
    "id",
    "blockId",
    "author",
    "createdAt",
    "message",
    "status",
  ])
    if (!(key in annotation))
      throw new AgentDocxError("INVALID_ARGUMENT", `${label} must have ${key}`);
  if (
    typeof annotation.id !== "string" ||
    !annotationIdPattern.test(annotation.id)
  )
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `${label}.id must be an annotation ID`,
    );
  if (
    typeof annotation.blockId !== "string" ||
    !blockIdPattern.test(annotation.blockId)
  )
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `${label}.blockId must be a block ID`,
    );
  if (annotation.range !== undefined) {
    const range = assertRecord(annotation.range, `${label}.range`);
    assertKeys(range, ["start", "end"], `${label}.range`);
    const start = requiredInteger(range, "start");
    const end = requiredInteger(range, "end");
    if (end < start)
      throw new AgentDocxError(
        "INVALID_ARGUMENT",
        `${label}.range.end must not precede start`,
      );
  }
  if (annotation.author !== null)
    assertChangeSetActor(annotation.author, `${label}.author`);
  if (annotation.createdAt !== null) {
    if (typeof annotation.createdAt !== "string")
      throw new AgentDocxError(
        "INVALID_ARGUMENT",
        `${label}.createdAt must be a date-time or null`,
      );
    assertDateTime(annotation.createdAt, `${label}.createdAt`);
  }
  requiredText(annotation, "message", label);
  if (annotation.status !== "open" && annotation.status !== "resolved")
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `${label}.status must be open or resolved`,
    );
};

export const assertResolutionDecisions = (value: unknown): void => {
  const decisions = assertRecord(value, "decisions");
  for (const [id, decision] of Object.entries(decisions)) {
    if (
      !changeIdPattern.test(id) ||
      (decision !== "accept" && decision !== "reject")
    )
      throw new AgentDocxError(
        "CHANGESET_INVALID",
        "Resolution decisions must map canonical change IDs to accept or reject",
      );
  }
};

export const assertReviewRange = (value: unknown): void => {
  const range = assertRecord(value, "range");
  assertKeys(range, ["start", "length"], "range");
  const start = requiredInteger(range, "start");
  const length = requiredInteger(range, "length");
  if (length < 0)
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      "Review range length must not be negative",
    );
  void start;
};

export const assertFilingSetId = (value: Record<string, unknown>): void => {
  if (typeof value.id !== "string" || !documentIdPattern.test(value.id))
    throw new AgentDocxError("INVALID_ARGUMENT", "id must be a filing set ID");
};

export const assertImportParams = (value: Record<string, unknown>): void => {
  assertKeys(
    value,
    [
      "input",
      "attachments",
      "inspectOnly",
      "documentId",
      "output",
      "author",
      "message",
    ],
    "docx.import params",
  );
  requiredString(value, "input");
  optionalString(value, "attachments");
  if (typeof value.inspectOnly !== "boolean")
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      "inspectOnly must be a boolean",
    );
  if (value.inspectOnly) {
    for (const key of ["documentId", "output", "author", "message"] as const)
      if (hasOwn(value, key))
        throw new AgentDocxError(
          "INVALID_ARGUMENT",
          "Inspect-only DOCX import is stateless",
        );
    return;
  }
  requiredDocumentId(value, "documentId");
  requiredString(value, "output");
  actor(value.author);
  requiredString(value, "message");
};

export const isStatelessAgentRequest = (
  action: AgentAction,
  params: Record<string, unknown>,
): boolean =>
  action === "docx.inspect" ||
  (action === "docx.import" && params.inspectOnly === true);

const assertAgentParams = (
  action: AgentAction,
  params: Record<string, unknown>,
): void => {
  switch (action) {
    case "project.init":
      assertProjectInput(params, false);
      return;
    case "project.add":
      assertProjectInput(params, true);
      return;
    case "project.get":
      noOptions(params, "project.get params");
      return;
    case "document.configure":
      assertKeys(
        params,
        ["documentId", "baseRevision", "changes", "author", "message"],
        "document.configure params",
      );
      requiredDocumentId(params, "documentId");
      asRevision(params.baseRevision, "baseRevision", true);
      assertConfigUpdate(params.changes);
      actor(params.author);
      requiredString(params, "message");
      return;
    case "document.get":
    case "document.validate":
    case "draft.guidance":
      assertKeys(params, ["documentId", "revision"], `${action} params`);
      requiredDocumentId(params, "documentId");
      optionalRevision(params, "revision");
      return;
    case "document.measure":
      assertKeys(
        params,
        ["documentId", "revision", "options"],
        "document.measure params",
      );
      if (params.options !== undefined)
        assertRendererOptions(params.options, "document.measure options", [
          "renderer",
          "officeTimeoutMs",
          "paragraphDiagnostics",
          "sectionDiagnostics",
          "lineDiagnostics",
          "trim",
          "word",
          "libreoffice",
        ]);
      return;
    case "revision.checkpoint":
      assertKeys(
        params,
        ["documentId", "baseRevision", "author", "message"],
        "revision.checkpoint params",
      );
      requiredDocumentId(params, "documentId");
      asRevision(params.baseRevision, "baseRevision", true);
      actor(params.author);
      requiredString(params, "message");
      return;
    case "revision.list":
      assertKeys(
        params,
        ["documentId", "limit", "cursor"],
        "revision.list params",
      );
      requiredDocumentId(params, "documentId");
      optionalInteger(params, "limit", 1, 1000);
      if (params.cursor !== undefined) requiredRevisionId(params, "cursor");
      return;
    case "revision.get":
      assertKeys(params, ["documentId", "revision"], "revision.get params");
      requiredDocumentId(params, "documentId");
      asRevision(params.revision, "revision");
      return;
    case "revision.restore":
      assertKeys(
        params,
        ["documentId", "baseRevision", "targetRevision", "author", "message"],
        "revision.restore params",
      );
      requiredDocumentId(params, "documentId");
      asRevision(params.baseRevision, "baseRevision");
      asRevision(params.targetRevision, "targetRevision");
      actor(params.author);
      requiredString(params, "message");
      return;
    case "revision.diff":
      assertKeys(
        params,
        ["documentId", "baseRevision", "headRevision", "output"],
        "revision.diff params",
      );
      requiredDocumentId(params, "documentId");
      asRevision(params.baseRevision, "baseRevision");
      asRevision(params.headRevision, "headRevision");
      optionalString(params, "output");
      return;
    case "revision.resolve":
      assertKeys(
        params,
        ["documentId", "changeSet", "decisions", "author", "message"],
        "revision.resolve params",
      );
      requiredDocumentId(params, "documentId");
      assertChangeSet(params.changeSet);
      assertResolutionDecisions(params.decisions);
      actor(params.author);
      requiredString(params, "message");
      return;
    case "draft.evaluate":
      assertKeys(params, ["patch", "renderer"], "draft.evaluate params");
      assertSourcePatch(params.patch);
      optionalEnum(params, "renderer", rendererModes);
      return;
    case "draft.apply":
      assertKeys(
        params,
        ["patch", "patchHash", "gate", "author", "message"],
        "draft.apply params",
      );
      assertSourcePatch(params.patch);
      requiredRevisionId(params, "patchHash");
      optionalEnum(params, "gate", ["report", "not-worse", "pass"]);
      actor(params.author);
      requiredString(params, "message");
      return;
    case "review.add":
      assertKeys(
        params,
        ["documentId", "revision", "blockId", "range", "author", "message"],
        "review.add params",
      );
      requiredDocumentId(params, "documentId");
      asRevision(params.revision, "revision");
      if (
        typeof params.blockId !== "string" ||
        !blockIdPattern.test(params.blockId)
      )
        throw new AgentDocxError(
          "INVALID_ARGUMENT",
          "blockId must be a block ID",
        );
      if (params.range !== undefined) assertReviewRange(params.range);
      actor(params.author);
      requiredString(params, "message");
      return;
    case "review.resolve":
      assertKeys(
        params,
        ["documentId", "revision", "annotationId", "author", "message"],
        "review.resolve params",
      );
      requiredDocumentId(params, "documentId");
      asRevision(params.revision, "revision");
      if (
        typeof params.annotationId !== "string" ||
        !annotationIdPattern.test(params.annotationId)
      )
        throw new AgentDocxError(
          "INVALID_ARGUMENT",
          "annotationId must be an annotation ID",
        );
      actor(params.author);
      requiredString(params, "message");
      return;
    case "docx.export": {
      assertKeys(
        params,
        ["documentId", "revision", "mode", "baseRevision", "output", "options"],
        "docx.export params",
      );
      requiredDocumentId(params, "documentId");
      asRevision(params.revision, "revision");
      const mode = requiredEnum(params, "mode", ["clean", "redline", "pdf"]);
      requiredString(params, "output");
      if (
        (mode === "clean" || mode === "pdf") &&
        hasOwn(params, "baseRevision")
      )
        throw new AgentDocxError(
          "INVALID_ARGUMENT",
          `${mode === "clean" ? "Clean" : "PDF"} DOCX export forbids baseRevision`,
        );
      if (mode === "redline") asRevision(params.baseRevision, "baseRevision");
      if (params.options !== undefined)
        assertRendererOptions(params.options, "docx.export options", [
          "renderer",
          "officeTimeoutMs",
          "word",
          "libreoffice",
        ]);
      return;
    }
    case "docx.importRedline":
      assertKeys(
        params,
        ["documentId", "input", "attachments", "author", "message"],
        "docx.importRedline params",
      );
      requiredDocumentId(params, "documentId");
      requiredString(params, "input");
      optionalString(params, "attachments");
      actor(params.author);
      requiredString(params, "message");
      return;
    case "filingSet.add":
      assertKeys(
        params,
        ["id", "label", "documentIds", "pageCap"],
        "filingSet.add params",
      );
      assertFilingSetId(params);
      optionalString(params, "label");
      if (
        !Array.isArray(params.documentIds) ||
        params.documentIds.length === 0 ||
        params.documentIds.some(
          (entry) =>
            typeof entry !== "string" || !documentIdPattern.test(entry),
        )
      )
        throw new AgentDocxError(
          "INVALID_ARGUMENT",
          "documentIds must be a non-empty array of document IDs",
        );
      const seen: Record<string, true> = {};
      for (const entry of params.documentIds as string[]) {
        if (seen[entry])
          throw new AgentDocxError(
            "INVALID_ARGUMENT",
            "documentIds must not contain duplicates",
          );
        seen[entry] = true;
      }
      optionalInteger(params, "pageCap", 1);
      return;
    case "filingSet.remove":
    case "filingSet.get":
    case "filingSet.validate":
      assertKeys(params, ["id"], `${action} params`);
      assertFilingSetId(params);
      return;
    case "docx.import":
      assertImportParams(params);
      return;
    case "docx.inspect":
      assertKeys(params, ["input"], "docx.inspect params");
      requiredString(params, "input");
      return;
  }
};

export const parseAgentRequest = (value: unknown): AgentRequest => {
  const request = assertRecord(value, "Agent request");
  assertKeys(
    request,
    ["schemaVersion", "id", "action", "project", "params"],
    "Agent request",
  );
  if (request.schemaVersion !== 1)
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      "Agent request schemaVersion must be 1",
    );
  if (
    hasOwn(request, "id") &&
    request.id !== null &&
    typeof request.id !== "string" &&
    (typeof request.id !== "number" || !Number.isFinite(request.id))
  )
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      "Agent request id must be a string, finite number, or null",
    );
  if (
    typeof request.action !== "string" ||
    !(agentActions as readonly string[]).includes(request.action)
  )
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      "Agent request action is invalid",
    );
  if (
    hasOwn(request, "project") &&
    (typeof request.project !== "string" || request.project === "")
  )
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      "Agent request project must be a non-empty path",
    );
  const action = request.action as AgentAction;
  const params = assertRecord(request.params, "Agent request params");
  assertAgentParams(action, params);
  if (hasOwn(request, "project") && isStatelessAgentRequest(action, params))
    throw new AgentDocxError(
      "INVALID_ARGUMENT",
      `${action} is stateless and forbids project`,
    );
  return {
    schemaVersion: 1,
    ...(hasOwn(request, "id") ? { id: request.id as AgentRequestId } : {}),
    action,
    ...(hasOwn(request, "project")
      ? { project: request.project as string }
      : {}),
    params: params as AgentParamsForAction<typeof action>,
  } as AgentRequest;
};
