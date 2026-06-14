import { defineTool, ToolError } from "@achmadya-dev/mcp-core";
import { z } from "zod";
import { Excel } from "../excel/excel.js";

const filePathField = z.string().describe("Absolute path to the local Excel file (.xlsx)");

const deleteSheetInputSchema = z.object({
  filePath: filePathField,
  sheetName: z.string(),
});

const operationResultOutputShape = z.object({
  success: z.boolean(),
  message: z.string(),
});

export const excel_delete_sheet = defineTool({
  name: "excel_delete_sheet",
  description: "Delete a worksheet (cannot delete the only sheet).",
  inputSchema: deleteSheetInputSchema,
  outputSchema: operationResultOutputShape,
  handler: async (args) => {
    try {
      const { filePath, sheetName } = args;
      const excel = await Excel.open(filePath);
      excel.removeSheet(sheetName);
      await excel.save();
      return { success: true as const, message: `Sheet '${sheetName}' deleted` };
    } catch (error) {
      const msg = error instanceof Error ? error.message : String(error);
      throw new ToolError("Failed to delete sheet: " + msg);
    }
  },
});
