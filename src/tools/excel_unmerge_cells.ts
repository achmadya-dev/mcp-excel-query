import { defineTool, ToolError } from "@achmadya-dev/mcp-core";
import { z } from "zod";
import { Excel } from "../excel/excel.js";

const filePathField = z.string().describe("Absolute path to the local Excel file (.xlsx)");

const sheetNameOptional = z
  .string()
  .optional()
  .describe("Name of the worksheet (defaults to the first sheet)");

const unmergeCellsInputSchema = z.object({
  filePath: filePathField,
  sheetName: sheetNameOptional,
  startCell: z.string(),
  endCell: z.string(),
});

const operationResultOutputShape = z.object({
  success: z.boolean(),
  message: z.string(),
});

export const excel_unmerge_cells = defineTool({
  name: "excel_unmerge_cells",
  description: "Unmerge a previously merged range.",
  inputSchema: unmergeCellsInputSchema,
  outputSchema: operationResultOutputShape,
  handler: async (args) => {
    try {
      const excel = await Excel.open(args.filePath);
      const sheet = excel.sheet(args.sheetName);
      const bounds = sheet.range(args.startCell, args.endCell);
      sheet.raw.unMergeCells(bounds.startRow, bounds.startCol, bounds.endRow, bounds.endCol);
      await excel.save();
      return {
        success: true as const,
        message: `Unmerged cells ${args.startCell}:${args.endCell}`,
      };
    } catch (error) {
      const msg = error instanceof Error ? error.message : String(error);
      throw new ToolError("Failed to unmerge cells: " + msg);
    }
  },
});
