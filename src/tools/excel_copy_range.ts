import { defineTool, ToolError } from "@achmadya-dev/mcp-core";
import { z } from "zod";
import { Excel } from "../excel/excel.js";
import { Range } from "../excel/range.js";

const filePathField = z.string().describe("Absolute path to the local Excel file (.xlsx)");

const sheetNameOptional = z
  .string()
  .optional()
  .describe("Name of the worksheet (defaults to the first sheet)");

const copyRangeInputSchema = z.object({
  filePath: filePathField,
  sheetName: sheetNameOptional,
  sourceStart: z.string(),
  sourceEnd: z.string(),
  targetStart: z.string(),
  targetSheet: z.string().optional(),
});

const operationResultOutputShape = z.object({
  success: z.boolean(),
  message: z.string(),
});

export const excel_copy_range = defineTool({
  name: "excel_copy_range",
  description: "Copy a cell range to another location (optionally another sheet).",
  inputSchema: copyRangeInputSchema,
  outputSchema: operationResultOutputShape,
  handler: async (args) => {
    try {
      const excel = await Excel.open(args.filePath);
      const sourceWs = excel.sheet(args.sheetName).raw;
      const targetWs = args.targetSheet ? excel.sheet(args.targetSheet).raw : sourceWs;

      const src = Range.between(args.sourceStart, args.sourceEnd);
      const tgt = Range.parse(args.targetStart);
      const rowOffset = tgt.startRow - src.startRow;
      const colOffset = tgt.startCol - src.startCol;

      for (let r = src.startRow; r <= src.endRow; r++) {
        for (let c = src.startCol; c <= src.endCol; c++) {
          const srcCell = sourceWs.getCell(r, c);
          const dstCell = targetWs.getCell(r + rowOffset, c + colOffset);
          dstCell.value = srcCell.value;
          dstCell.style = Object.assign({}, srcCell.style);
        }
      }

      await excel.save();
      return {
        success: true as const,
        message: `Copied ${args.sourceStart}:${args.sourceEnd} to ${args.targetStart} on sheet '${targetWs.name}'`,
      };
    } catch (error) {
      const msg = error instanceof Error ? error.message : String(error);
      throw new ToolError("Failed to copy range: " + msg);
    }
  },
});
