import { defineTool, ToolError } from "@achmadya-dev/mcp-core";
import { z } from "zod";
import { Excel } from "../excel/excel.js";
import { Range } from "../excel/range.js";

const filePathField = z.string().describe("Absolute path to the local Excel file (.xlsx)");

const sheetNameOptional = z
  .string()
  .optional()
  .describe("Name of the worksheet (defaults to the first sheet)");

const setDimensionsInputSchema = z.object({
  filePath: filePathField,
  sheetName: sheetNameOptional,
  range: z.string().optional(),
  startCell: z.string().optional(),
  endCell: z.string().optional(),
  rows: z.array(z.number()).optional().describe("1-based row numbers for rowHeight"),
  columns: z
    .array(z.union([z.number(), z.string()]))
    .optional()
    .describe("1-based column indexes or letters for columnWidth"),
  rowHeight: z.number().optional().describe("Row height in points"),
  columnWidth: z.number().optional().describe("Column width in Excel character units"),
  defaultRowHeight: z.number().optional().describe("Sheet default row height"),
  defaultColumnWidth: z.number().optional().describe("Sheet default column width"),
});

const operationResultOutputShape = z.object({
  success: z.boolean(),
  message: z.string(),
});

function columnIndex(column: number | string): number {
  if (typeof column === "number") {
    if (!Number.isInteger(column) || column < 1) {
      throw new Error(`Invalid column index: ${column}`);
    }
    return column;
  }
  return Range.letterToCol(column);
}

function collectTargetRows(params: {
  rows?: number[];
  range?: string;
  startCell?: string;
  endCell?: string;
}): number[] {
  if (params.rows && params.rows.length > 0) {
    return params.rows;
  }

  if (params.range) {
    const bounds = Range.parse(params.range);
    return Array.from(
      { length: bounds.endRow - bounds.startRow + 1 },
      (_, i) => bounds.startRow + i
    );
  }

  if (params.startCell) {
    const bounds = Range.between(params.startCell, params.endCell);
    return Array.from(
      { length: bounds.endRow - bounds.startRow + 1 },
      (_, i) => bounds.startRow + i
    );
  }

  return [];
}

function collectTargetColumns(params: {
  columns?: Array<number | string>;
  range?: string;
  startCell?: string;
  endCell?: string;
}): number[] {
  if (params.columns && params.columns.length > 0) {
    return params.columns.map(columnIndex);
  }

  if (params.range) {
    const bounds = Range.parse(params.range);
    return Array.from(
      { length: bounds.endCol - bounds.startCol + 1 },
      (_, i) => bounds.startCol + i
    );
  }

  if (params.startCell) {
    const bounds = Range.between(params.startCell, params.endCell);
    return Array.from(
      { length: bounds.endCol - bounds.startCol + 1 },
      (_, i) => bounds.startCol + i
    );
  }

  return [];
}

export const excel_set_dimensions = defineTool({
  name: "excel_set_dimensions",
  description:
    "Adjust row heights and column widths by range, explicit rows/columns, or sheet defaults.",
  inputSchema: setDimensionsInputSchema,
  outputSchema: operationResultOutputShape,
  handler: async (args) => {
    try {
      const { filePath, sheetName, rowHeight, columnWidth, defaultRowHeight, defaultColumnWidth } =
        args;

      if (
        rowHeight === undefined &&
        columnWidth === undefined &&
        defaultRowHeight === undefined &&
        defaultColumnWidth === undefined
      ) {
        throw new Error(
          "Provide rowHeight, columnWidth, defaultRowHeight, and/or defaultColumnWidth"
        );
      }

      const excel = await Excel.open(filePath);
      const sheet = excel.sheet(sheetName);
      const worksheet = sheet.raw;

      const targetRows = collectTargetRows(args);
      const targetColumns = collectTargetColumns(args);

      if (rowHeight !== undefined) {
        if (targetRows.length === 0) {
          throw new Error("rowHeight requires rows, range, or startCell");
        }
        for (const rowNumber of targetRows) {
          worksheet.getRow(rowNumber).height = rowHeight;
        }
      }

      if (columnWidth !== undefined) {
        if (targetColumns.length === 0) {
          throw new Error("columnWidth requires columns, range, or startCell");
        }
        for (const colNumber of targetColumns) {
          worksheet.getColumn(colNumber).width = columnWidth;
        }
      }

      if (defaultRowHeight !== undefined) {
        worksheet.properties.defaultRowHeight = defaultRowHeight;
      }

      if (defaultColumnWidth !== undefined) {
        worksheet.properties.defaultColWidth = defaultColumnWidth;
      }

      await excel.save();

      const parts: string[] = [];
      if (rowHeight !== undefined) {
        parts.push(`row height ${rowHeight} on rows ${targetRows.join(", ")}`);
      }
      if (columnWidth !== undefined) {
        const labels = targetColumns.map((c) => Range.colToLetter(c));
        parts.push(`column width ${columnWidth} on columns ${labels.join(", ")}`);
      }
      if (defaultRowHeight !== undefined) {
        parts.push(`default row height ${defaultRowHeight}`);
      }
      if (defaultColumnWidth !== undefined) {
        parts.push(`default column width ${defaultColumnWidth}`);
      }

      return {
        success: true as const,
        message: `Updated ${parts.join("; ")} on sheet '${worksheet.name}'`,
      };
    } catch (error) {
      const msg = error instanceof Error ? error.message : String(error);
      throw new ToolError("Failed to set dimensions: " + msg);
    }
  },
});
