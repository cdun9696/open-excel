import type { AgentTool } from "@mariozechner/pi-agent-core";
import type { Static, TObject } from "@sinclair/typebox";

export interface ToolResult {
  content: { type: "text"; text: string }[];
  details: undefined;
}

export function defineTool<T extends TObject>(config: {
  name: string;
  label: string;
  description: string;
  parameters: T;
  execute: (toolCallId: string, params: Static<T>, signal?: AbortSignal) => Promise<ToolResult>;
}): AgentTool {
  return config as unknown as AgentTool;
}

export function toolSuccess(data: unknown): ToolResult {
  return {
    content: [{ type: "text", text: JSON.stringify(data) }],
    details: undefined,
  };
}

export function toolError(message: string): ToolResult {
  return {
    content: [{ type: "text", text: JSON.stringify({ success: false, error: message }) }],
    details: undefined,
  };
}
