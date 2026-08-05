import type { JsonValue } from "../types.js";
import type {
  Actor,
  AddressableBlock,
  BlockId,
  LeafAddressableBlock,
  ReviewAnnotation,
  RevisionId,
} from "../legal/model.js";
import type { ValidationResult } from "../legal/rules.js";
import type { SerializableProjectMeasurementResult } from "../project/contracts.js";

export type { ReviewAnnotation } from "../legal/model.js";

export type RevisionRecord = {
  schemaVersion: 1;
  id: RevisionId;
  documentId: string;
  parents: readonly RevisionId[];
  createdAt: string;
  author: Actor;
  message: string;
  sourceObject: RevisionId;
  documentConfigObject: RevisionId;
  dependencyObjects: Readonly<Record<string, RevisionId>>;
  workingTreeHash: RevisionId;
  legalDocumentObject: RevisionId;
  annotationsObject: RevisionId;
  deltaObject?: RevisionId;
  resolutionObject?: RevisionId;
  tool: { name: "agent-docx"; version: string; schemaVersion: 1 };
};

export type AnnotationChangeBase = { id: `c_${string}` };
export type AnnotationChange =
  | (AnnotationChangeBase & { kind: "add"; newValue: ReviewAnnotation })
  | (AnnotationChangeBase & {
      kind: "replace";
      oldValue: ReviewAnnotation;
      newValue: ReviewAnnotation;
    })
  | (AnnotationChangeBase & { kind: "remove"; oldValue: ReviewAnnotation });

export type ChangeAttribution = {
  author: Actor | null;
  createdAt: string | null;
  sourceRevisionId?: string;
};

export type AttributionSpan = {
  start: number;
  end: number;
  attribution: ChangeAttribution;
};

export type BlockLocation = {
  collection: "body" | "footnotes";
  parentId: BlockId | null;
  index: number;
  sourceOffset: number;
};

export type ContainerShell = {
  blockId: BlockId;
  kind: "list" | "exhibit" | "length-exclusion";
  attributes: JsonValue;
  sourceRanges: readonly { start: number; end: number; text: string }[];
};

export type ChangeBase = {
  id: `c_${string}`;
  attribution: ChangeAttribution;
};

export type Change =
  | (ChangeBase & {
      kind: "insert-block";
      blockId: BlockId;
      to: BlockLocation;
      newSource: { start: number; end: number; text: string };
      block: AddressableBlock;
      newAttributionSpans: readonly AttributionSpan[];
    })
  | (ChangeBase & {
      kind: "delete-block";
      blockId: BlockId;
      from: BlockLocation;
      oldSource: { start: number; end: number; text: string };
      oldBlock: AddressableBlock;
      oldAttributionSpans: readonly AttributionSpan[];
    })
  | (ChangeBase & {
      kind: "move-block";
      blockId: BlockId;
      from: BlockLocation;
      to: BlockLocation;
      oldSource: { start: number; end: number; text: string };
      newSource: { start: number; end: number; text: string };
      oldAttributionSpans: readonly AttributionSpan[];
      newAttributionSpans: readonly AttributionSpan[];
    })
  | (ChangeBase & {
      kind: "replace-block";
      blockId: BlockId;
      oldBlock: LeafAddressableBlock;
      newBlock: LeafAddressableBlock;
      oldAttributionSpans: readonly AttributionSpan[];
      newAttributionSpans: readonly AttributionSpan[];
    })
  | (ChangeBase & {
      kind: "replace-container-shell";
      blockId: BlockId;
      oldShell: ContainerShell;
      newShell: ContainerShell;
      oldAttributionSpans: readonly AttributionSpan[];
      newAttributionSpans: readonly AttributionSpan[];
    })
  | (ChangeBase & {
      kind: "insert-text";
      blockId: BlockId;
      oldOffset: number;
      newSource: { start: number; end: number; text: string };
      newText: string;
      newAttributionSpans: readonly AttributionSpan[];
    })
  | (ChangeBase & {
      kind: "delete-text";
      blockId: BlockId;
      oldSource: { start: number; end: number; text: string };
      newOffset: number;
      oldText: string;
      oldAttributionSpans: readonly AttributionSpan[];
    })
  | (ChangeBase & {
      kind: "replace-text";
      blockId: BlockId;
      oldSource: { start: number; end: number; text: string };
      newSource: { start: number; end: number; text: string };
      oldText: string;
      newText: string;
      oldAttributionSpans: readonly AttributionSpan[];
      newAttributionSpans: readonly AttributionSpan[];
    })
  | (ChangeBase & { kind: "add-config"; path: string; newValue: JsonValue })
  | (ChangeBase & { kind: "remove-config"; path: string; oldValue: JsonValue })
  | (ChangeBase & {
      kind: "replace-config";
      path: string;
      oldValue: JsonValue;
      newValue: JsonValue;
    })
  | (ChangeBase & { kind: "add-dependency"; key: string; newObject: RevisionId })
  | (ChangeBase & { kind: "remove-dependency"; key: string; oldObject: RevisionId })
  | (ChangeBase & {
      kind: "replace-dependency";
      key: string;
      oldObject: RevisionId;
      newObject: RevisionId;
    });

export type ChangeSet = {
  schemaVersion: 1;
  id: RevisionId;
  documentId: string;
  baseRevision: RevisionId;
  headRevision: RevisionId;
  changes: readonly Change[];
  annotations: readonly AnnotationChange[];
};

export type RevisionDeltaRecord = {
  schemaVersion: 1;
  parentSourceObject: RevisionId;
  parentDocumentConfigObject: RevisionId;
  changes: readonly Change[];
  annotations: readonly AnnotationChange[];
};

export type ResolutionRecord = {
  schemaVersion: 1;
  changeSet: ChangeSet;
  decisions: Readonly<Record<`c_${string}`, "accept" | "reject">>;
};

export type RevisionPage = {
  schemaVersion: 1;
  items: readonly RevisionRecord[];
  nextCursor: RevisionId | null;
};

export type RevisionMutationResult = {
  schemaVersion: 1;
  revision: RevisionRecord;
  head: RevisionId;
  sourceSha256: RevisionId;
  workingTreeHash: RevisionId;
  measurement: SerializableProjectMeasurementResult;
  validation: ValidationResult;
};

