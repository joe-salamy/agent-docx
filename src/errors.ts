import { AgentDocxError, type ErrorCode, type JsonValue } from "./types.js";

export type ErrorPayload = {
  code: ErrorCode;
  message: string;
  details?: Record<string, JsonValue>;
};

/** One projection from any thrown value to the protocol error shape. */
export const toErrorPayload = (error: unknown): ErrorPayload => {
  if (error instanceof AgentDocxError)
    return {
      code: error.code,
      message: error.message,
      ...(error.details ? { details: error.details } : {}),
    };
  return {
    code: "INTERNAL_ERROR",
    message: error instanceof Error ? error.message : String(error),
  };
};
