import { AgentDocxError, type ErrorCode } from "./types.js";

/** Validates one JSON object boundary before callers inspect named fields. */
export const objectRecord = (
  value: unknown,
  label: string,
  code: ErrorCode = "INVALID_ARGUMENT",
): Record<string, unknown> => {
  if (typeof value !== "object" || value === null || Array.isArray(value))
    throw new AgentDocxError(code, `${label} must be an object`);
  return value as Record<string, unknown>;
};
