import { z } from "zod";

// Base schemas
export const cellValueSchema = z.union([z.string(), z.number(), z.boolean(), z.null()]);

// Tools Input Schemas
export const listSheetsInputSchema = {
  filePath: z.string().describe("Absolute path to the local Excel file (.xlsx)"),
};

export const readSheetInputSchema = {
  filePath: z.string().describe("Absolute path to the local Excel file (.xlsx)"),
  sheetName: z.string().optional().describe("Name of the worksheet to read (defaults to the first sheet)"),
  range: z.string().optional().describe("Cell range in A1 notation (e.g., 'A1:C10')"),
  headerRow: z.number().optional().describe("Row number containing headers, 1-indexed (defaults to 1). Set to 0 if there are no headers."),
  limit: z.number().optional().describe("Maximum number of data rows to return (for pagination)"),
  offset: z.number().optional().describe("Number of data rows to skip (for pagination)"),
};

export const getMetadataInputSchema = {
  filePath: z.string().describe("Absolute path to the local Excel file (.xlsx)"),
};

export const createFileInputSchema = {
  filePath: z.string().describe("Absolute path where the new Excel file (.xlsx) will be created"),
  sheetName: z.string().optional().describe("Name of the first sheet (defaults to 'Sheet1')"),
  headers: z.array(z.string()).optional().describe("Optional list of header names to write in the first row"),
};

export const createSheetInputSchema = {
  filePath: z.string().describe("Absolute path to the local Excel file (.xlsx)"),
  sheetName: z.string().describe("Name of the new worksheet to create"),
};

const borderSideSchema = z.object({
  style: z.string().describe("Border style (e.g. 'thin', 'medium', 'thick', 'double')"),
  color: z.string().optional().describe("HEX color code for border (e.g. '000000')"),
}).optional();

export const cellStyleSchema = z.object({
  font: z.object({
    name: z.string().optional().describe("Font name (e.g., 'Arial')"),
    size: z.number().optional().describe("Font size in points (e.g., 12)"),
    bold: z.boolean().optional().describe("Make text bold"),
    italic: z.boolean().optional().describe("Make text italic"),
    underline: z.boolean().optional().describe("Underline text"),
    color: z.string().optional().describe("HEX color code for text (e.g., 'FF0000' for red)"),
  }).optional().describe("Text font settings"),
  fill: z.object({
    type: z.enum(["pattern", "gradient"]).optional().describe("Fill type (defaults to 'pattern')"),
    pattern: z.string().optional().describe("Pattern type (defaults to 'solid')"),
    fgColor: z.string().optional().describe("HEX background/foreground color of the cell (e.g., 'FFFF00' for yellow)"),
    bgColor: z.string().optional().describe("HEX pattern background color"),
  }).optional().describe("Cell background fill settings"),
  alignment: z.object({
    horizontal: z.enum(["left", "center", "right", "justify", "distributed"]).optional().describe("Horizontal alignment"),
    vertical: z.enum(["top", "middle", "bottom", "justify", "distributed"]).optional().describe("Vertical alignment"),
    wrapText: z.boolean().optional().describe("Wrap text in cell"),
  }).optional().describe("Alignment settings"),
  border: z.object({
    top: borderSideSchema,
    left: borderSideSchema,
    bottom: borderSideSchema,
    right: borderSideSchema,
  }).optional().describe("Cell borders"),
});

export const writeCellInputSchema = {
  filePath: z.string().describe("Absolute path to the local Excel file (.xlsx)"),
  sheetName: z.string().optional().describe("Name of the worksheet (defaults to the first sheet)"),
  cell: z.string().describe("Cell address in A1 notation (e.g., 'B5')"),
  value: cellValueSchema.optional().describe("Value to write into the cell (can be string, number, boolean, or null)"),
  style: cellStyleSchema.optional().describe("Optional styling to apply to the cell (font, fill/color, border, alignment)"),
};

export const appendRowsInputSchema = {
  filePath: z.string().describe("Absolute path to the local Excel file (.xlsx)"),
  sheetName: z.string().optional().describe("Name of the worksheet (defaults to the first sheet)"),
  rows: z.array(z.array(cellValueSchema)).describe("Array of rows to append, where each row is an array of cell values"),
};

// Response Output Shapes
export const listSheetsOutputShape = {
  sheets: z.array(z.string()),
};

export const readSheetOutputShape = {
  sheetName: z.string(),
  headers: z.array(z.string()).nullable(),
  rows: z.array(z.record(z.string(), z.any())),
  totalRows: z.number(),
};

export const getMetadataOutputShape = {
  filePath: z.string(),
  creator: z.string().nullable(),
  lastModifiedBy: z.string().nullable(),
  created: z.string().nullable(),
  modified: z.string().nullable(),
  sheets: z.array(
    z.object({
      name: z.string(),
      rowCount: z.number(),
      columnCount: z.number(),
    })
  ),
};

export const operationResultOutputShape = {
  success: z.boolean(),
  message: z.string(),
};
