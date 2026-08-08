import { objectRecord } from "../json-contract.js";
import { AgentDocxError } from "../types.js";

export const unsupported = (message: string): never => {
  throw new AgentDocxError("DOCX_IMPORT_UNSUPPORTED", message);
};

export const asObject = (
  value: unknown,
  label: string,
): Record<string, unknown> =>
  objectRecord(value, label, { code: "DOCX_IMPORT_UNSUPPORTED" });

export const codePointCompare = (left: string, right: string): number =>
  left < right ? -1 : left > right ? 1 : 0;
