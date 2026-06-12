import { defineTool, ToolError } from "./server.js";
import * as excel from "./excel/excel.js";
import * as schema from "./excel/schema.js";

function wrap<A extends object, R>(
  label: string,
  fn: (args: A) => Promise<R>
): (args: A) => Promise<R> {
  return async (args: A) => {
    try {
      return await fn(args);
    } catch (error) {
      const msg = error instanceof Error ? error.message : String(error);
      throw new ToolError(`${label}: ${msg}`);
    }
  };
}

export const excel_read_sheet = defineTool({
  name: "excel_read_sheet",
  description:
    "Read worksheet data as JSON rows. Supports range, pagination, showFormula, showStyle, dateFormat (or dateFormat: cell for per-cell numFmt). Returns mergedCells and per-cell validation metadata when present.",
  inputSchema: schema.readSheetInputSchema,
  outputSchema: schema.readSheetOutputShape,
  handler: wrap("Failed to read sheet", excel.readSheet),
});

export const excel_get_metadata = defineTool({
  name: "excel_get_metadata",
  description:
    "Workbook metadata: sheet names, dimensions, creator, dates. Use includeRanges for usedRange and mergedCells per sheet (replaces list_sheets).",
  inputSchema: schema.getMetadataInputSchema,
  outputSchema: schema.getMetadataOutputShape,
  handler: wrap("Failed to get metadata", excel.getMetadata),
});

export const excel_create_file = defineTool({
  name: "excel_create_file",
  description: "Create a new Excel file with optional first sheet and headers.",
  inputSchema: schema.createFileInputSchema,
  outputSchema: schema.operationResultOutputShape,
  handler: wrap("Failed to create Excel file", excel.createFile),
});

export const excel_write_range = defineTool({
  name: "excel_write_range",
  description:
    "Write data: values (2D), data (objects), append rows (append:true), create blank sheet (newSheet:true), single cell/formula (startCell + values), optional style for one cell.",
  inputSchema: schema.writeRangeInputSchema,
  outputSchema: schema.operationResultOutputShape,
  handler: wrap("Failed to write range", excel.writeRange),
});

export const excel_format_range = defineTool({
  name: "excel_format_range",
  description:
    "Format a cell range (font, fill, border, alignment, numFmt). Set mergeCells:true to merge. Use for styling after excel_write_range.",
  inputSchema: schema.formatRangeInputSchema,
  outputSchema: schema.operationResultOutputShape,
  handler: wrap("Failed to format range", excel.formatRange),
});

export const excel_copy_sheet = defineTool({
  name: "excel_copy_sheet",
  description: "Copy an existing worksheet to a new sheet name.",
  inputSchema: schema.copySheetInputSchema,
  outputSchema: schema.operationResultOutputShape,
  handler: wrap("Failed to copy sheet", excel.copySheet),
});

export const excel_rename_sheet = defineTool({
  name: "excel_rename_sheet",
  description: "Rename a worksheet.",
  inputSchema: schema.renameSheetInputSchema,
  outputSchema: schema.operationResultOutputShape,
  handler: wrap("Failed to rename sheet", excel.renameSheet),
});

export const excel_delete_sheet = defineTool({
  name: "excel_delete_sheet",
  description: "Delete a worksheet (cannot delete the only sheet).",
  inputSchema: schema.deleteSheetInputSchema,
  outputSchema: schema.operationResultOutputShape,
  handler: wrap("Failed to delete sheet", excel.deleteSheet),
});

export const excel_copy_range = defineTool({
  name: "excel_copy_range",
  description: "Copy a cell range to another location (optionally another sheet).",
  inputSchema: schema.copyRangeInputSchema,
  outputSchema: schema.operationResultOutputShape,
  handler: wrap("Failed to copy range", excel.copyRange),
});

export const excel_delete_range = defineTool({
  name: "excel_delete_range",
  description:
    "Delete a cell range and shift remaining cells. Use shiftDirection 'up' for row deletion, 'left' for column deletion.",
  inputSchema: schema.deleteRangeInputSchema,
  outputSchema: schema.operationResultOutputShape,
  handler: wrap("Failed to delete range", excel.deleteRange),
});

export const excel_unmerge_cells = defineTool({
  name: "excel_unmerge_cells",
  description: "Unmerge a previously merged range.",
  inputSchema: schema.unmergeCellsInputSchema,
  outputSchema: schema.operationResultOutputShape,
  handler: wrap("Failed to unmerge cells", excel.unmergeCells),
});

export const excel_create_table = defineTool({
  name: "excel_create_table",
  description: "Create a native Excel table from a data range.",
  inputSchema: schema.createTableInputSchema,
  outputSchema: schema.operationResultOutputShape,
  handler: wrap("Failed to create table", excel.createTable),
});

export const excel_insert_rows = defineTool({
  name: "excel_insert_rows",
  description: "Insert empty rows at a 1-based row index.",
  inputSchema: schema.insertRowsInputSchema,
  outputSchema: schema.operationResultOutputShape,
  handler: wrap("Failed to insert rows", excel.insertRows),
});

export const excel_insert_columns = defineTool({
  name: "excel_insert_columns",
  description: "Insert empty columns at a 1-based column index.",
  inputSchema: schema.insertColumnsInputSchema,
  outputSchema: schema.operationResultOutputShape,
  handler: wrap("Failed to insert columns", excel.insertColumns),
});

export const excel_set_sheet_visibility = defineTool({
  name: "excel_set_sheet_visibility",
  description: "Show, hide, or very-hide a worksheet (state: visible, hidden, veryHidden).",
  inputSchema: schema.setSheetVisibilityInputSchema,
  outputSchema: schema.operationResultOutputShape,
  handler: wrap("Failed to set sheet visibility", excel.setSheetVisibility),
});

export const excel_set_data_validation = defineTool({
  name: "excel_set_data_validation",
  description:
    "Set or clear data validation on a range. Supports list, whole, decimal, date, textLength, and custom rules with formulae.",
  inputSchema: schema.setDataValidationInputSchema,
  outputSchema: schema.operationResultOutputShape,
  handler: wrap("Failed to set data validation", excel.setDataValidation),
});

export const excel_set_dimensions = defineTool({
  name: "excel_set_dimensions",
  description:
    "Adjust row heights and column widths by range, explicit rows/columns, or sheet defaults.",
  inputSchema: schema.setDimensionsInputSchema,
  outputSchema: schema.operationResultOutputShape,
  handler: wrap("Failed to set dimensions", excel.setDimensions),
});

export const allExcelTools = [
  excel_read_sheet,
  excel_get_metadata,
  excel_create_file,
  excel_write_range,
  excel_format_range,
  excel_copy_sheet,
  excel_rename_sheet,
  excel_delete_sheet,
  excel_copy_range,
  excel_delete_range,
  excel_unmerge_cells,
  excel_create_table,
  excel_insert_rows,
  excel_insert_columns,
  excel_set_sheet_visibility,
  excel_set_data_validation,
  excel_set_dimensions,
] as const;
