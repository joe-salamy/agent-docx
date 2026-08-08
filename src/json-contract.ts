import { AgentDocxError, type ErrorCode } from "./types.js";

export type ObjectRecordOptions = {
  code?: ErrorCode;
  message?: string;
};

/** Validates one JSON object boundary before callers inspect named fields. */
export const objectRecord = (
  value: unknown,
  label: string,
  options: ObjectRecordOptions = {},
): Record<string, unknown> => {
  if (typeof value !== "object" || value === null || Array.isArray(value))
    throw new AgentDocxError(
      options.code ?? "INVALID_ARGUMENT",
      options.message ?? `${label} must be an object`,
    );
  return value as Record<string, unknown>;
};

type DefinedProps<T> = {
  [K in keyof T as T[K] extends undefined ? never : K]: Exclude<
    T[K],
    undefined
  >;
};

/**
 * Drops undefined-valued entries so the result can be spread into
 * `exactOptionalPropertyTypes` targets without weakening their types.
 */
export const definedProps = <T extends object>(value: T): DefinedProps<T> => {
  const out: Record<string, unknown> = {};
  for (const [key, entry] of Object.entries(value))
    if (entry !== undefined) out[key] = entry;
  return out as DefinedProps<T>;
};

/** True when every property of `value` is in `keys` (unknown properties rejected). */
export const hasOnlyKeys = (
  value: Record<string, unknown>,
  keys: readonly string[],
): boolean => {
  const allowed = new Set(keys);
  return Object.keys(value).every((key) => allowed.has(key));
};

/**
 * True when `value`'s property set is exactly `keys`: `hasOnlyKeys` plus equal
 * cardinality, for records where every allowed key is required.
 */
export const hasExactKeys = (
  value: Record<string, unknown>,
  keys: readonly string[],
): boolean =>
  hasOnlyKeys(value, keys) && Object.keys(value).length === keys.length;
