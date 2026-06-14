import { defineTool, ToolError } from "@achmadya-dev/mcp-core";
import { z } from "zod";
import { Excel } from "../excel/excel.js";

const filePathField = z.string().describe("Absolute path to the local Excel file (.xlsx)");

const renameSheetInputSchema = z.object({
  filePath: filePathField,
  oldName: z.string(),
  newName: z.string(),
});

const operationResultOutputShape = z.object({
  success: z.boolean(),
  message: z.string(),
});

export const excel_rename_sheet = defineTool({
  name: "excel_rename_sheet",
  description: "Rename a worksheet.",
  inputSchema: renameSheetInputSchema,
  outputSchema: operationResultOutputShape,
  handler: async (args) => {
    try {
      const { filePath, oldName, newName } = args;
      const excel = await Excel.open(filePath);
      const sheet = excel.sheet(oldName);

      if (excel.hasSheet(newName)) {
        throw new Error(`Worksheet already exists: "${newName}"`);
      }

      sheet.raw.name = newName;
      await excel.save();
      return { success: true as const, message: `Sheet renamed from '${oldName}' to '${newName}'` };
    } catch (error) {
      const msg = error instanceof Error ? error.message : String(error);
      throw new ToolError("Failed to rename sheet: " + msg);
    }
  },
});
