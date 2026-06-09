import { z } from "zod";

export const cellValueSchema = z.union([z.string(), z.number(), z.boolean(), z.null()]);

const filePathField = z.string().describe("Absolute path to the local Excel file (.xlsx)");
const sheetNameOptional = z.string().optional().describe("Name of the worksheet (defaults to the first sheet)");

export const readSheetInputSchema = {
  filePath: filePathField,
  sheetName: sheetNameOptional,
  range: z.string().optional().describe("Cell range in A1 notation (e.g., 'A1:C10')"),
  startCell: z.string().optional(),
  endCell: z.string().optional(),
  headerRow: z.number().optional().describe("Row number containing headers, 1-indexed (defaults to 1). Set to 0 if there are no headers."),
  limit: z.number().optional(),
  offset: z.number().optional(),
  showFormula: z.boolean().optional(),
  showStyle: z.boolean().optional(),
  previewOnly: z.boolean().optional(),
  dateFormat: z
    .string()
    .optional()
    .describe(
      'Format date cells for output. Use Excel-style tokens (e.g. "dd mmm yyyy") or "cell" to use each cell numFmt. Defaults to ISO when omitted.'
    ),
};

export const getMetadataInputSchema = {
  filePath: filePathField,
  includeRanges: z.boolean().optional().describe("Include usedRange and mergedCells per sheet"),
};

export const createFileInputSchema = {
  filePath: z.string().describe("Absolute path where the new Excel file (.xlsx) will be created"),
  sheetName: z.string().optional().describe("Name of the first sheet (defaults to 'Sheet1')"),
  headers: z.array(z.string()).optional(),
};

const borderSideSchema = z.object({
  style: z.string().optional(),
  color: z.string().optional(),
}).optional();

export const cellStyleSchema = z.object({
  font: z.object({
    name: z.string().optional(),
    size: z.number().optional(),
    bold: z.boolean().optional(),
    italic: z.boolean().optional(),
    underline: z.boolean().optional(),
    color: z.string().optional(),
  }).optional(),
  fill: z.object({
    type: z.enum(["pattern", "gradient"]).optional(),
    pattern: z.string().optional(),
    fgColor: z.string().optional(),
    bgColor: z.string().optional(),
  }).optional(),
  alignment: z.object({
    horizontal: z.enum(["left", "center", "right", "justify", "distributed"]).optional(),
    vertical: z.enum(["top", "middle", "bottom", "justify", "distributed"]).optional(),
    wrapText: z.boolean().optional(),
  }).optional(),
  border: z.object({
    top: borderSideSchema,
    left: borderSideSchema,
    bottom: borderSideSchema,
    right: borderSideSchema,
  }).optional(),
  numFmt: z.string().optional(),
});

export const writeRangeInputSchema = {
  filePath: filePathField,
  sheetName: sheetNameOptional,
  newSheet: z.boolean().optional().describe("Create a new sheet (blank if no values/data, or write initial data)"),
  append: z.boolean().optional().describe("Append rows to the bottom of the sheet instead of writing at startCell"),
  range: z.string().optional(),
  startCell: z.string().optional().describe("Start cell (default A1); use for single-cell writes including formulas (=...)"),
  values: z.array(z.array(cellValueSchema)).optional(),
  data: z.array(z.record(z.string(), z.any())).optional(),
  style: cellStyleSchema.optional().describe("Optional style when writing a single cell via startCell + one value"),
};

export const formatRangeInputSchema = {
  filePath: filePathField,
  sheetName: sheetNameOptional,
  startCell: z.string(),
  endCell: z.string().optional(),
  range: z.string().optional(),
  bold: z.boolean().optional(),
  italic: z.boolean().optional(),
  underline: z.boolean().optional(),
  fontSize: z.number().optional(),
  fontColor: z.string().optional(),
  bgColor: z.string().optional(),
  borderStyle: z.string().optional(),
  borderColor: z.string().optional(),
  numberFormat: z.string().optional(),
  alignment: z.string().optional(),
  wrapText: z.boolean().optional(),
  mergeCells: z.boolean().optional().describe("Merge the range while applying format"),
  styles: z.array(z.array(cellStyleSchema.nullable())).optional(),
};

export const copySheetInputSchema = {
  filePath: filePathField,
  sourceSheet: z.string(),
  targetSheet: z.string(),
};

export const renameSheetInputSchema = {
  filePath: filePathField,
  oldName: z.string(),
  newName: z.string(),
};

export const deleteSheetInputSchema = {
  filePath: filePathField,
  sheetName: z.string(),
};

export const copyRangeInputSchema = {
  filePath: filePathField,
  sheetName: sheetNameOptional,
  sourceStart: z.string(),
  sourceEnd: z.string(),
  targetStart: z.string(),
  targetSheet: z.string().optional(),
};

export const deleteRangeInputSchema = {
  filePath: filePathField,
  sheetName: sheetNameOptional,
  startCell: z.string(),
  endCell: z.string(),
  shiftDirection: z.enum(["up", "left"]).optional().describe("Use 'up' to delete full rows, 'left' to delete full columns"),
};

export const unmergeCellsInputSchema = {
  filePath: filePathField,
  sheetName: sheetNameOptional,
  startCell: z.string(),
  endCell: z.string(),
};

export const createTableInputSchema = {
  filePath: filePathField,
  sheetName: sheetNameOptional,
  dataRange: z.string(),
  tableName: z.string().optional(),
  tableStyle: z.string().optional(),
};

export const insertRowsInputSchema = {
  filePath: filePathField,
  sheetName: sheetNameOptional,
  startRow: z.number(),
  count: z.number().optional(),
};

export const insertColumnsInputSchema = {
  filePath: filePathField,
  sheetName: sheetNameOptional,
  startCol: z.number(),
  count: z.number().optional(),
};

export const readSheetOutputShape = {
  sheetName: z.string(),
  headers: z.array(z.string()).nullable(),
  rows: z.array(z.record(z.string(), z.any())),
  totalRows: z.number(),
  mergedCells: z.array(z.string()).optional(),
  previewOnly: z.boolean().optional(),
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
      usedRange: z.string().nullable().optional(),
      mergedCells: z.array(z.string()).optional(),
    })
  ),
};

export const operationResultOutputShape = {
  success: z.boolean(),
  message: z.string(),
};
