import { defineTool, ToolError } from "@achmadya-dev/mcp-core";
import { z } from "zod";
import { Excel } from "../excel/excel.js";

const createFileInputSchema = z.object({
  filePath: z.string().describe("Absolute path where the new Excel file (.xlsx) will be created"),
  sheetName: z.string().optional().describe("Name of the first sheet (defaults to 'Sheet1')"),
  headers: z.array(z.string()).optional(),
});

const operationResultOutputShape = z.object({
  success: z.boolean(),
  message: z.string(),
});

export const excel_create_file = defineTool({
  name: "excel_create_file",
  description: "Create a new Excel file with optional first sheet and headers.",
  inputSchema: createFileInputSchema,
  outputSchema: operationResultOutputShape,
  handler: async (args) => {
    try {
      const { filePath, sheetName = "Sheet1", headers } = args;
      const excel = Excel.create(filePath, { sheetName, headers });
      await excel.save();
      return { success: true as const, message: `Excel file created successfully at ${filePath}` };
    } catch (error) {
      const msg = error instanceof Error ? error.message : String(error);
      throw new ToolError("Failed to create Excel file: " + msg);
    }
  },
});
