import ExcelJS from "exceljs";
import type { JsonValue } from "../types.js";
import {
  extractCellStyle,
  getPagingCellsLimit,
  loadWorkbook,
  numToColLetter,
  parseA1Range,
  readCellDisplayValue,
  resolveRangeBounds,
  resolveWorksheet,
} from "./utils.js";

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

export async function readSheet(params: {
  filePath: string;
  sheetName?: string;
  range?: string;
  startCell?: string;
  endCell?: string;
  headerRow?: number;
  limit?: number;
  offset?: number;
  showFormula?: boolean;
  showStyle?: boolean;
  previewOnly?: boolean;
  dateFormat?: string;
}) {
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
  } = params;

  const rangeStr = buildRangeFromParams(params.range, params.startCell, params.endCell);
  const workbook = await loadWorkbook(filePath);
  const worksheet = resolveWorksheet(workbook, sheetName);

  let startRow = 1;
  let endRow = worksheet.rowCount;
  let startCol = 1;
  let endCol = worksheet.actualColumnCount || 1;

  if (rangeStr) {
    const parsedRange = parseA1Range(rangeStr);
    startRow = parsedRange.startRow;
    endRow = parsedRange.endRow;
    startCol = parsedRange.startCol;
    endCol = parsedRange.endCol;
  }

  const cellCount = (endRow - startRow + 1) * (endCol - startCol + 1);
  const pagingLimit = getPagingCellsLimit();
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
      const cellVal = readCellDisplayValue(row.getCell(c).value, showFormula);
      headers.push(
        cellVal !== null && cellVal !== undefined && cellVal !== ""
          ? String(cellVal)
          : numToColLetter(c)
      );
    }
  } else {
    for (let c = startCol; c <= endCol; c++) {
      headers.push(numToColLetter(c));
    }
  }

  const allRows: Record<string, JsonValue>[] = [];
  const startDataRow = headerRowIdx > 0 ? Math.max(startRow, headerRowIdx + 1) : startRow;
  const effectiveEndRow = previewOnly
    ? Math.min(endRow, startDataRow + 9)
    : endRow;

  for (let r = startDataRow; r <= effectiveEndRow; r++) {
    const row = worksheet.getRow(r);
    const rowData: Record<string, JsonValue> = {};
    let hasValue = false;

    for (let c = startCol; c <= endCol; c++) {
      const excelCell = row.getCell(c);
      const headerKey = headers[c - startCol];
      const cellVal = readCellDisplayValue(excelCell.value, showFormula, {
        dateFormat,
        cellNumFmt: excelCell.numFmt,
      });
      rowData[headerKey] = cellVal;

      if (showStyle) {
        const style = extractCellStyle(excelCell);
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
}
