import { defineTool, ToolError } from "@achmadya-dev/mcp-core";
import { z } from "zod";
import { Excel } from "../excel/excel.js";

const filePathField = z.string().describe("Absolute path to the local Excel file (.xlsx)");

const sheetNameOptional = z
  .string()
  .optional()
  .describe("Name of the worksheet (defaults to the first sheet)");

const insertColumnsInputSchema = z.object({
  filePath: filePathField,
  sheetName: sheetNameOptional,
  startCol: z.number(),
  count: z.number().optional(),
});

const operationResultOutputShape = z.object({
  success: z.boolean(),
  message: z.string(),
});

export const excel_insert_columns = defineTool({
  name: "excel_insert_columns",
  description: "Insert empty columns at a 1-based column index.",
  inputSchema: insertColumnsInputSchema,
  outputSchema: operationResultOutputShape,
  handler: async (args) => {
    try {
      const { filePath, sheetName, startCol, count = 1 } = args;
      const excel = await Excel.open(filePath);
      const sheet = excel.sheet(sheetName);

      sheet.raw.eachRow((row) => {
        row.splice(startCol, 0, ...Array(count).fill(null));
      });

      await excel.save();
      return {
        success: true as const,
        message: `Inserted ${count} column(s) at column ${startCol}`,
      };
    } catch (error) {
      const msg = error instanceof Error ? error.message : String(error);
      throw new ToolError("Failed to insert columns: " + msg);
    }
  },
});
