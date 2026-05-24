import { defineTool, ToolError } from "./server.js";
import {
  listSheets,
  readSheet,
  getMetadata,
  createFile,
  createSheet,
  writeCell,
  appendRows,
} from "./excel/excel.js";
import {
  listSheetsInputSchema,
  listSheetsOutputShape,
  readSheetInputSchema,
  readSheetOutputShape,
  getMetadataInputSchema,
  getMetadataOutputShape,
  createFileInputSchema,
  createSheetInputSchema,
  writeCellInputSchema,
  appendRowsInputSchema,
  operationResultOutputShape,
} from "./excel/schema.js";

export const excel_list_sheets = defineTool({
  name: "excel_list_sheets",
  description: "List the names of all worksheets inside an Excel file (.xlsx)",
  inputSchema: listSheetsInputSchema,
  outputSchema: listSheetsOutputShape,
  handler: async ({ filePath }) => {
    try {
      return await listSheets(filePath);
    } catch (error) {
      const msg = error instanceof Error ? error.message : String(error);
      throw new ToolError(`Failed to list sheets: ${msg}`);
    }
  },
});

export const excel_read_sheet = defineTool({
  name: "excel_read_sheet",
  description: "Read data rows from a specific worksheet in an Excel file. Supports range (A1 notation), headerRow index, and paginated reading (limit/offset).",
  inputSchema: readSheetInputSchema,
  outputSchema: readSheetOutputShape,
  handler: async (args) => {
    try {
      return await readSheet(args);
    } catch (error) {
      const msg = error instanceof Error ? error.message : String(error);
      throw new ToolError(`Failed to read sheet: ${msg}`);
    }
  },
});

export const excel_get_metadata = defineTool({
  name: "excel_get_metadata",
  description: "Retrieve metadata from an Excel file, including creator, modification dates, and lists of sheets with their sizes.",
  inputSchema: getMetadataInputSchema,
  outputSchema: getMetadataOutputShape,
  handler: async ({ filePath }) => {
    try {
      return await getMetadata(filePath);
    } catch (error) {
      const msg = error instanceof Error ? error.message : String(error);
      throw new ToolError(`Failed to get metadata: ${msg}`);
    }
  },
});

export const excel_create_file = defineTool({
  name: "excel_create_file",
  description: "Create a new Excel file (.xlsx) with an optional first worksheet and optional headers.",
  inputSchema: createFileInputSchema,
  outputSchema: operationResultOutputShape,
  handler: async (args) => {
    try {
      return await createFile(args);
    } catch (error) {
      const msg = error instanceof Error ? error.message : String(error);
      throw new ToolError(`Failed to create Excel file: ${msg}`);
    }
  },
});

export const excel_create_sheet = defineTool({
  name: "excel_create_sheet",
  description: "Add a new blank worksheet to an existing Excel file (.xlsx).",
  inputSchema: createSheetInputSchema,
  outputSchema: operationResultOutputShape,
  handler: async (args) => {
    try {
      return await createSheet(args);
    } catch (error) {
      const msg = error instanceof Error ? error.message : String(error);
      throw new ToolError(`Failed to create worksheet: ${msg}`);
    }
  },
});

export const excel_write_cell = defineTool({
  name: "excel_write_cell",
  description: "Write or update a value in a specific cell (e.g., 'A1') of a worksheet. The cell value can be a string, number, boolean, or null.",
  inputSchema: writeCellInputSchema,
  outputSchema: operationResultOutputShape,
  handler: async (args) => {
    try {
      return await writeCell(args);
    } catch (error) {
      const msg = error instanceof Error ? error.message : String(error);
      throw new ToolError(`Failed to write cell: ${msg}`);
    }
  },
});

export const excel_append_rows = defineTool({
  name: "excel_append_rows",
  description: "Append one or more rows of data to the bottom of a worksheet. Each row is represented as an array of cell values.",
  inputSchema: appendRowsInputSchema,
  outputSchema: operationResultOutputShape,
  handler: async (args) => {
    try {
      return await appendRows(args);
    } catch (error) {
      const msg = error instanceof Error ? error.message : String(error);
      throw new ToolError(`Failed to append rows: ${msg}`);
    }
  },
});
