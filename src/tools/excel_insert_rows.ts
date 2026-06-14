import { defineTool, ToolError } from "@achmadya-dev/mcp-core";
import { z } from "zod";
import { Excel } from "../excel/excel.js";

const filePathField = z.string().describe("Absolute path to the local Excel file (.xlsx)");

const sheetNameOptional = z
  .string()
  .optional()
  .describe("Name of the worksheet (defaults to the first sheet)");

const insertRowsInputSchema = z.object({
  filePath: filePathField,
  sheetName: sheetNameOptional,
  startRow: z.number(),
  count: z.number().optional(),
});

const operationResultOutputShape = z.object({
  success: z.boolean(),
  message: z.string(),
});

export const excel_insert_rows = defineTool({
  name: "excel_insert_rows",
  description: "Insert empty rows at a 1-based row index.",
  inputSchema: insertRowsInputSchema,
  outputSchema: operationResultOutputShape,
  handler: async (args) => {
    try {
      const { filePath, sheetName, startRow, count = 1 } = args;
      const excel = await Excel.open(filePath);
      const sheet = excel.sheet(sheetName);
      sheet.raw.spliceRows(startRow, 0, ...Array.from({ length: count }, () => []));
      await excel.save();
      return { success: true as const, message: `Inserted ${count} row(s) at row ${startRow}` };
    } catch (error) {
      const msg = error instanceof Error ? error.message : String(error);
      throw new ToolError("Failed to insert rows: " + msg);
    }
  },
});
