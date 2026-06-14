import { defineTool, ToolError } from "@achmadya-dev/mcp-core";
import { z } from "zod";
import { Excel } from "../excel/excel.js";

const filePathField = z.string().describe("Absolute path to the local Excel file (.xlsx)");

const sheetNameOptional = z
  .string()
  .optional()
  .describe("Name of the worksheet (defaults to the first sheet)");

const deleteRangeInputSchema = z.object({
  filePath: filePathField,
  sheetName: sheetNameOptional,
  startCell: z.string(),
  endCell: z.string(),
  shiftDirection: z
    .enum(["up", "left"])
    .optional()
    .describe("Use 'up' to delete full rows, 'left' to delete full columns"),
});

const operationResultOutputShape = z.object({
  success: z.boolean(),
  message: z.string(),
});

export const excel_delete_range = defineTool({
  name: "excel_delete_range",
  description:
    "Delete a cell range and shift remaining cells. Use shiftDirection 'up' for row deletion, 'left' for column deletion.",
  inputSchema: deleteRangeInputSchema,
  outputSchema: operationResultOutputShape,
  handler: async (args) => {
    try {
      const excel = await Excel.open(args.filePath);
      const sheet = excel.sheet(args.sheetName);
      const bounds = sheet.range(args.startCell, args.endCell);
      const shift = args.shiftDirection ?? "up";

      if (shift === "up") {
        sheet.raw.spliceRows(bounds.startRow, bounds.endRow - bounds.startRow + 1);
      } else {
        for (let r = bounds.startRow; r <= bounds.endRow; r++) {
          const row = sheet.raw.getRow(r);
          row.splice(bounds.startCol, bounds.endCol - bounds.startCol + 1);
        }
      }

      await excel.save();
      return {
        success: true as const,
        message: `Deleted range ${args.startCell}:${args.endCell} (shift ${shift})`,
      };
    } catch (error) {
      const msg = error instanceof Error ? error.message : String(error);
      throw new ToolError("Failed to delete range: " + msg);
    }
  },
});
