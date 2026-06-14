import { defineTool, ToolError } from "@achmadya-dev/mcp-core";
import { z } from "zod";
import { Cell } from "../excel/cell.js";
import { Excel } from "../excel/excel.js";
import { Range } from "../excel/range.js";

const filePathField = z.string().describe("Absolute path to the local Excel file (.xlsx)");

const sheetNameOptional = z
  .string()
  .optional()
  .describe("Name of the worksheet (defaults to the first sheet)");

const cellValueSchema = z.union([z.string(), z.number(), z.boolean(), z.null()]);

const borderSideSchema = z
  .object({
    style: z.string().optional(),
    color: z.string().optional(),
  })
  .optional();
const cellStyleSchema = z.object({
  font: z
    .object({
      name: z.string().optional(),
      size: z.number().optional(),
      bold: z.boolean().optional(),
      italic: z.boolean().optional(),
      underline: z.boolean().optional(),
      color: z.string().optional(),
    })
    .optional(),
  fill: z
    .object({
      type: z.enum(["pattern", "gradient"]).optional(),
      pattern: z.string().optional(),
      fgColor: z.string().optional(),
      bgColor: z.string().optional(),
    })
    .optional(),
  alignment: z
    .object({
      horizontal: z.enum(["left", "center", "right", "justify", "distributed"]).optional(),
      vertical: z.enum(["top", "middle", "bottom", "justify", "distributed"]).optional(),
      wrapText: z.boolean().optional(),
    })
    .optional(),
  border: z
    .object({
      top: borderSideSchema,
      left: borderSideSchema,
      bottom: borderSideSchema,
      right: borderSideSchema,
    })
    .optional(),
  numFmt: z.string().optional(),
});

const writeRangeInputSchema = z.object({
  filePath: filePathField,
  sheetName: sheetNameOptional,
  newSheet: z
    .boolean()
    .optional()
    .describe("Create a new sheet (blank if no values/data, or write initial data)"),
  append: z
    .boolean()
    .optional()
    .describe("Append rows to the bottom of the sheet instead of writing at startCell"),
  range: z.string().optional(),
  startCell: z
    .string()
    .optional()
    .describe("Start cell (default A1); use for single-cell writes including formulas (=...)"),
  values: z.array(z.array(cellValueSchema)).optional(),
  data: z.array(z.record(z.string(), z.any())).optional(),
  style: cellStyleSchema
    .optional()
    .describe("Optional style when writing a single cell via startCell + one value"),
});

const operationResultOutputShape = z.object({
  success: z.boolean(),
  message: z.string(),
});

function recordsToGrid(data: Record<string, unknown>[]): unknown[][] {
  if (data.length === 0) return [];
  const keys = Object.keys(data[0]);
  const grid: unknown[][] = [keys];
  for (const record of data) {
    grid.push(keys.map((k) => record[k] ?? null));
  }
  return grid;
}

export const excel_write_range = defineTool({
  name: "excel_write_range",
  description:
    "Write data: values (2D), data (objects), append rows (append:true), create blank sheet (newSheet:true), single cell/formula (startCell + values), optional style for one cell.",
  inputSchema: writeRangeInputSchema,
  outputSchema: operationResultOutputShape,
  handler: async (args) => {
    try {
      const {
        filePath,
        sheetName,
        newSheet = false,
        append = false,
        range,
        startCell = "A1",
        style,
      } = args;

      let grid: unknown[][] | undefined = args.values;
      if (!grid && args.data) {
        grid = recordsToGrid(args.data);
      }

      const excel = await Excel.open(filePath);
      let sheet;

      if (newSheet) {
        if (!sheetName) throw new Error("sheetName is required when newSheet is true");
        sheet = excel.addSheet(sheetName);

        if (!grid || grid.length === 0) {
          await excel.save();
          return { success: true as const, message: `Sheet '${sheetName}' created in ${filePath}` };
        }
      } else {
        sheet = excel.sheet(sheetName);
      }

      const worksheet = sheet.raw;

      if (!grid || grid.length === 0) {
        throw new Error(
          "values, data, or append rows are required unless newSheet creates a blank sheet"
        );
      }

      if (append) {
        for (const rowData of grid) {
          worksheet.addRow(rowData);
        }
        await excel.save();
        return {
          success: true as const,
          message: `Appended ${grid.length} row(s) to sheet '${worksheet.name}'`,
        };
      }

      const origin = range ? Range.parse(range).startRow : Range.parse(startCell).startRow;
      const originCol = range ? Range.parse(range).startCol : Range.parse(startCell).startCol;

      for (let r = 0; r < grid.length; r++) {
        const row = grid[r];
        for (let c = 0; c < row.length; c++) {
          const cell = worksheet.getCell(origin + r, originCol + c);
          Cell.set(cell, row[c]);
          if (style && grid.length === 1 && row.length === 1) {
            Cell.applyStyle(cell, style);
          }
        }
      }

      await excel.save();
      const targetRange =
        range ??
        `${startCell}:${Range.address(origin + grid.length - 1, originCol + grid[0].length - 1)}`;
      return {
        success: true as const,
        message: `Wrote ${grid.length} row(s) to sheet '${worksheet.name}' at ${targetRange}`,
      };
    } catch (error) {
      const msg = error instanceof Error ? error.message : String(error);
      throw new ToolError("Failed to write range: " + msg);
    }
  },
});
