import { defineTool, ToolError } from "@achmadya-dev/mcp-core";
import { z } from "zod";
import type ExcelJS from "exceljs";
import type { JsonValue } from "@achmadya-dev/mcp-core";
import { Cell } from "../excel/cell.js";
import { Excel } from "../excel/excel.js";
import { Range } from "../excel/range.js";

const filePathField = z.string().describe("Absolute path to the local Excel file (.xlsx)");

const sheetNameOptional = z
  .string()
  .optional()
  .describe("Name of the worksheet (defaults to the first sheet)");

const readSheetInputSchema = z.object({
  filePath: filePathField,
  sheetName: sheetNameOptional,
  range: z.string().optional().describe("Cell range in A1 notation (e.g., 'A1:C10')"),
  startCell: z.string().optional(),
  endCell: z.string().optional(),
  headerRow: z
    .number()
    .optional()
    .describe(
      "Row number containing headers, 1-indexed (defaults to 1). Set to 0 if there are no headers."
    ),
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
});

const readSheetOutputShape = z.object({
  sheetName: z.string(),
  headers: z.array(z.string()).nullable(),
  rows: z.array(z.record(z.string(), z.any())),
  totalRows: z.number(),
  mergedCells: z.array(z.string()).optional(),
  previewOnly: z.boolean().optional(),
});

function buildRangeFromParams(
  range?: string,
  startCell?: string,
  endCell?: string
): string | undefined {
  if (range) return range;
  if (startCell && endCell) return `${startCell}:${endCell}`;
  if (startCell) return startCell;
  return undefined;
}

export const excel_read_sheet = defineTool({
  name: "excel_read_sheet",
  description:
    "Read worksheet data as JSON rows. Supports range, pagination, showFormula, showStyle, dateFormat (or dateFormat: cell for per-cell numFmt). Returns mergedCells and per-cell validation metadata when present.",
  inputSchema: readSheetInputSchema,
  outputSchema: readSheetOutputShape,
  handler: async (args) => {
    try {
      const {
        filePath,
        sheetName,
        headerRow = 1,
        limit,
        offset,
        showFormula = false,
        showStyle = false,
        previewOnly = false,
        dateFormat,
      } = args;

      const rangeStr = buildRangeFromParams(args.range, args.startCell, args.endCell);
      const excel = await Excel.open(filePath);
      const sheet = excel.sheet(sheetName);
      const worksheet = sheet.raw;

      let startRow = 1;
      let endRow = worksheet.rowCount;
      let startCol = 1;
      let endCol = worksheet.actualColumnCount || 1;

      if (rangeStr) {
        const parsedRange = Range.parse(rangeStr);
        startRow = parsedRange.startRow;
        endRow = parsedRange.endRow;
        startCol = parsedRange.startCol;
        endCol = parsedRange.endCol;
      }

      const cellCount = (endRow - startRow + 1) * (endCol - startCol + 1);
      const pagingLimit = Excel.pagingLimit();
      if (cellCount > pagingLimit && !previewOnly && limit === undefined) {
        throw new Error(
          `Range exceeds paging limit (${cellCount} cells > ${pagingLimit}). ` +
            `Use limit/offset, previewOnly, or set EXCEL_PAGING_CELLS_LIMIT.`
        );
      }

      const headers: string[] = [];
      const headerRowIdx = headerRow;

      if (headerRowIdx > 0 && headerRowIdx <= worksheet.rowCount) {
        const row = worksheet.getRow(headerRowIdx);
        for (let c = startCol; c <= endCol; c++) {
          const cellVal = Cell.readDisplay(row.getCell(c).value, showFormula);
          headers.push(
            cellVal !== null && cellVal !== undefined && cellVal !== ""
              ? String(cellVal)
              : Range.colToLetter(c)
          );
        }
      } else {
        for (let c = startCol; c <= endCol; c++) {
          headers.push(Range.colToLetter(c));
        }
      }

      const allRows: Record<string, JsonValue>[] = [];
      const startDataRow = headerRowIdx > 0 ? Math.max(startRow, headerRowIdx + 1) : startRow;
      const effectiveEndRow = previewOnly ? Math.min(endRow, startDataRow + 9) : endRow;

      for (let r = startDataRow; r <= effectiveEndRow; r++) {
        const row = worksheet.getRow(r);
        const rowData: Record<string, JsonValue> = {};
        let hasValue = false;

        for (let c = startCol; c <= endCol; c++) {
          const excelCell = row.getCell(c);
          const headerKey = headers[c - startCol];
          const cellVal = Cell.readDisplay(excelCell.value, showFormula, {
            dateFormat,
            cellNumFmt: excelCell.numFmt,
          });
          rowData[headerKey] = cellVal;

          if (showStyle) {
            const style = Cell.extractStyle(excelCell);
            if (style) {
              rowData[`__style_${headerKey}`] = JSON.parse(JSON.stringify(style)) as JsonValue;
            }
          }

          const validation = (excelCell as ExcelJS.Cell & { dataValidation?: unknown })
            .dataValidation;
          if (validation) {
            rowData[`__validation_${headerKey}`] = JSON.parse(
              JSON.stringify(validation)
            ) as JsonValue;
          }

          if (cellVal !== null && cellVal !== undefined && cellVal !== "") {
            hasValue = true;
          }
        }

        if (hasValue || rangeStr) {
          allRows.push(rowData);
        }
      }

      const totalRows = allRows.length;
      const off = offset ?? 0;
      const lim = limit ?? totalRows;
      const paginatedRows = allRows.slice(off, off + lim);

      const mergedCells = (worksheet.model.merges ?? []).map((m) => String(m));

      return {
        sheetName: worksheet.name,
        headers: headerRowIdx > 0 ? headers : null,
        rows: paginatedRows,
        totalRows,
        mergedCells,
        previewOnly,
      };
    } catch (error) {
      const msg = error instanceof Error ? error.message : String(error);
      throw new ToolError("Failed to read sheet: " + msg);
    }
  },
});
