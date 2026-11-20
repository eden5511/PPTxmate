import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import type { Static, TObject } from "@sinclair/typebox";
import { Value } from "@sinclair/typebox/value";

export type ToolResult = AgentToolResult<undefined>;

interface ToolConfig<T extends TObject> {
  name: string;
  label: string;
  description: string;
  parameters: T;
  execute: (
    toolCallId: string,
    params: Static<T>,
    signal?: AbortSignal,
  ) => Promise<ToolResult>;
}

export function defineTool<T extends TObject>(
  config: ToolConfig<T>,
): AgentTool {
  return {
    ...config,
    execute: async (toolCallId, params, signal) => {
      if (signal?.aborted) {
        throw new DOMException("Tool execution aborted", "AbortError");
      }
      if (!Value.Check(config.parameters, params)) {
        const fields = [...Value.Errors(config.parameters, params)]
          .slice(0, 5)
          .map((error) => error.path || "/");
        return toolError(
          `INVALID_ARGUMENTS: ${config.name}; check fields: ${fields.join(", ")}`,
        );
      }
      const result = await config.execute(toolCallId, params, signal);
      if (!result || !Array.isArray(result.content) || !result.content.length) {
        return toolError(
          "TOOL_RESULT_MISSING: execution returned no content; completion is unknown. Inspect state before retrying writes.",
        );
      }
      return result;
    },
  } as AgentTool;
}

export function toolSuccess(data: unknown): ToolResult {
  const result =
    typeof data === "object" && data !== null ? { ...data } : { result: data };
  return {
    content: [{ type: "text", text: JSON.stringify(result) }],
    details: undefined,
  };
}

export function toolError(message: string): ToolResult {
  return {
    content: [
      {
        type: "text",
        text: JSON.stringify({ success: false, error: message }),
      },
    ],
    details: undefined,
  };
}

export function toolText(text: string): ToolResult {
  return {
    content: [{ type: "text", text }],
    details: undefined,
  };
}
