import { Type } from "@sinclair/typebox";
import { describe, expect, it, vi } from "vitest";
import { defineTool, toolText, type ToolResult } from "../src/tools/types";

describe("tool contracts", () => {
  it("rejects missing parameters before execution", async () => {
    const execute = vi.fn(async () => toolText("ok"));
    const tool = defineTool({ name: "read", label: "Read", description: "test", parameters: Type.Object({ path: Type.String({ minLength: 1 }) }), execute });
    const result = await tool.execute("call", {});
    expect(JSON.stringify(result)).toContain("INVALID_ARGUMENTS");
    expect(execute).not.toHaveBeenCalled();
  });

  it("reports missing content instead of returning an empty envelope", async () => {
    const tool = defineTool({ name: "read", label: "Read", description: "test", parameters: Type.Object({}), execute: async () => ({ content: [], details: undefined }) as ToolResult });
    expect(JSON.stringify(await tool.execute("call", {}))).toContain("TOOL_RESULT_MISSING");
  });
});
