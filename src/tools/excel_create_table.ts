import { defineTool, ToolError } from "@achmadya-dev/mcp-core";
import { z } from "zod";
import { Excel } from "../excel/excel.js";
import { Range } from "../excel/range.js";

const filePathField = z.string().describe("Absolute path to the local Excel file (.xlsx)");

const sheetNameOptional = z
  .string()
  .optional()
  .describe("Name of the worksheet (defaults to the first sheet)");

const createTableInputSchema = z.object({
  filePath: filePathField,
  sheetName: sheetNameOptional,
  dataRange: z.string(),
  tableName: z.string().optional(),
  tableStyle: z.string().optional(),
});

const operationResultOutputShape = z.object({
  success: z.boolean(),
  message: z.string(),
});

export const excel_create_table = defineTool({
  name: "excel_create_table",
  description: "Create a native Excel table from a data range.",
  inputSchema: createTableInputSchema,
  outputSchema: operationResultOutputShape,
  handler: async (args) => {
    try {
      const excel = await Excel.open(args.filePath);
      const sheet = excel.sheet(args.sheetName);
      const worksheet = sheet.raw;
      const bounds = Range.parse(args.dataRange);

      const headerRow = worksheet.getRow(bounds.startRow);
      const columns: { name: string; filterButton?: boolean }[] = [];
      for (let c = bounds.startCol; c <= bounds.endCol; c++) {
        const val = headerRow.getCell(c).value;
        const name =
          val !== null && val !== undefined && val !== ""
            ? String(val)
            : `Column${c - bounds.startCol + 1}`;
        columns.push({ name, filterButton: true });
      }

      const rows: unknown[][] = [];
      for (let r = bounds.startRow + 1; r <= bounds.endRow; r++) {
        const row = worksheet.getRow(r);
        const rowData: unknown[] = [];
        for (let c = bounds.startCol; c <= bounds.endCol; c++) {
          rowData.push(row.getCell(c).value ?? null);
        }
        rows.push(rowData);
      }

      worksheet.addTable({
        name: args.tableName ?? `Table_${Date.now()}`,
        ref: args.dataRange,
        headerRow: true,
        style: {
          theme: (args.tableStyle ?? "TableStyleMedium9") as "TableStyleMedium9",
          showRowStripes: true,
        },
        columns,
        rows,
      });

      await excel.save();
      return {
        success: true as const,
        message: `Table '${args.tableName ?? "created"}' added on sheet '${worksheet.name}' at ${args.dataRange}`,
      };
    } catch (error) {
      const msg = error instanceof Error ? error.message : String(error);
      throw new ToolError("Failed to create table: " + msg);
    }
  },
});
