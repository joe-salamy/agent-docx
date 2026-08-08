import { dispatchAgentRequest } from "./agent-dispatch.js";
import {
  agentActions,
  isStatelessAgentRequest,
  parseAgentRequest,
} from "./agent-protocol.js";
import { serializeAgentValue } from "./agent-serialize.js";
import type { AgentDispatchResult } from "./agent-protocol.js";

export {
  agentActions,
  dispatchAgentRequest,
  isStatelessAgentRequest,
  parseAgentRequest,
  serializeAgentValue,
};
export type {
  AgentAction,
  AgentDispatchResult,
  AgentRequest,
} from "./agent-protocol.js";

/** Parses and executes one agent request. */
export const executeAgentRequest = async (
  raw: unknown,
  cwd = process.cwd(),
): Promise<AgentDispatchResult> =>
  dispatchAgentRequest(parseAgentRequest(raw), cwd);
