import type { AgentContext } from "@office-agents/core";
import { sandboxedEval } from "@office-agents/core";
import { Type } from "@sinclair/typebox";
import JSZip from "jszip";
import { centeredAxisBounds, sharedAxisBounds } from "../pptx/chart-axis";
import { chartSnapshot, validateChartEdits } from "../pptx/chart-validation";
import {
  slideTargetParameterProperties,
  toSlideMutationNotStartedError,
  toSlideTargetReference,
} from "../pptx/slide-directory";
import { safeRun, withSlideZip } from "../pptx/slide-zip";
import { escapeXml } from "../pptx/xml-utils";
import { unpackSlideZipResult } from "./slide-target-result";
import { defineTool, toolError, toolSuccess } from "./types";

/* global PowerPoint */

export function createEditSlideChartTool(ctx: AgentContext) {
  return defineTool({
    name: "edit_slide_chart",
    label: "Edit Slide Chart",
    description:
      "Add or edit charts in a PowerPoint slide by manipulating raw OOXML. " +
      "Always use this for data visualizations — never approximate charts with geometric shapes.",
    parameters: Type.Object({
      ...slideTargetParameterProperties,
      require_editable_data: Type.Optional(
        Type.Boolean({
          description: "Require an embedded Excel workbook for every chart. Set true when the user needs Edit Data.",
        }),
      ),
      code: Type.String({
        description:
          "Async function body receiving { zip, markDirty }. zip is a JSZip archive of the slide. " +
          "Call markDirty() if you modified files. " +
          "Globals: loadZip(bytes) for embedded XLSX, sharedAxisBounds(values, step?, includeZero?), centeredAxisBounds(values, minimumSpan?), escapeXml(text) for safe XML text embedding, " +
          "readFile(path) returns Promise<string> and readFileBuffer(path) returns Promise<Uint8Array> " +
          "to read files from the virtual filesystem (e.g. uploaded images, SVGs). " +
          "writeFile(path, content) returns Promise<void> to write string or Uint8Array to the virtual filesystem.",
      }),
      explanation: Type.Optional(
        Type.String({
          description: "Brief description (max 50 chars)",
          maxLength: 50,
        }),
      ),
    }),
    execute: async (_toolCallId, params) => {
      try {
        const target = toSlideTargetReference(params);
        const zipValue = await safeRun(async (context) => {
          return withSlideZip(context, target, async (args) => {
            const before = await chartSnapshot(args.zip);
            let dirty = false;
            const result = await sandboxedEval(params.code, {
              ...args,
              markDirty: () => { dirty = true; },
              escapeXml,
              loadZip: (data: Uint8Array) => JSZip.loadAsync(data),
              sharedAxisBounds,
              centeredAxisBounds,
              readFile: (path: string) => ctx.readFile(path),
              readFileBuffer: (path: string) => ctx.readFileBuffer(path),
              writeFile: (path: string, content: string | Uint8Array) =>
                ctx.writeFile(path, content),
              DOMParser,
              XMLSerializer,
            });
            if (dirty) {
              await validateChartEdits(args.zip, before, params.require_editable_data);
              args.markDirty();
            }
            return result;
          });
        });
        const { result, metadata } = unpackSlideZipResult(zipValue, target);

        return toolSuccess({
          success: true,
          ...metadata,
          result: result !== undefined ? result : null,
        });
      } catch (error) {
        const normalized = toSlideMutationNotStartedError(
          error,
          "Failed to edit slide chart",
        );
        return toolError(normalized.message, normalized);
      }
    },
    modifiedSlide: (params, result) =>
      result && typeof result.slideId === "string"
        ? result.slideId
        : (params.slide_id ?? params.slide_index),
  });
}
