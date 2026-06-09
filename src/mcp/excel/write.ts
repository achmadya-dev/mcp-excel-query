import ExcelJS from "exceljs";
import {
  applyStyleToCell,
  cellAddress,
  loadWorkbook,
  operationSuccess,
  parseA1Range,
  resolveWorksheet,
  saveWorkbook,
  setCellValue,
} from "./utils.js";

function recordsToGrid(data: Record<string, unknown>[]): unknown[][] {
  if (data.length === 0) return [];
  const keys = Object.keys(data[0]);
  const grid: unknown[][] = [keys];
  for (const record of data) {
    grid.push(keys.map((k) => record[k] ?? null));
  }
  return grid;
}

export async function writeRange(params: {
  filePath: string;
  sheetName?: string;
  newSheet?: boolean;
  append?: boolean;
  range?: string;
  startCell?: string;
  values?: unknown[][];
  data?: Record<string, unknown>[];
  style?: Record<string, unknown>;
}) {
  const { filePath, sheetName, newSheet = false, append = false, range, startCell = "A1", style } = params;

  let grid = params.values;
  if (!grid && params.data) {
    grid = recordsToGrid(params.data);
  }

  const workbook = await loadWorkbook(filePath);
  let worksheet: ExcelJS.Worksheet;

  if (newSheet) {
    if (!sheetName) throw new Error("sheetName is required when newSheet is true");
    if (workbook.getWorksheet(sheetName)) {
      throw new Error(`Worksheet already exists: "${sheetName}"`);
    }
    worksheet = workbook.addWorksheet(sheetName);

    if (!grid || grid.length === 0) {
      await saveWorkbook(workbook, filePath);
      return operationSuccess(`Sheet '${sheetName}' created in ${filePath}`);
    }
  } else {
    worksheet = resolveWorksheet(workbook, sheetName);
  }

  if (!grid || grid.length === 0) {
    throw new Error("values, data, or append rows are required unless newSheet creates a blank sheet");
  }

  if (append) {
    for (const rowData of grid) {
      worksheet.addRow(rowData);
    }
    await saveWorkbook(workbook, filePath);
    return operationSuccess(
      `Appended ${grid.length} row(s) to sheet '${worksheet.name}'`
    );
  }

  const origin = range ? parseA1Range(range).startRow : parseA1Range(startCell).startRow;
  const originCol = range ? parseA1Range(range).startCol : parseA1Range(startCell).startCol;

  for (let r = 0; r < grid.length; r++) {
    const row = grid[r];
    for (let c = 0; c < row.length; c++) {
      const cell = worksheet.getCell(origin + r, originCol + c);
      setCellValue(cell, row[c]);
      if (style && grid.length === 1 && row.length === 1) {
        applyStyleToCell(cell, style);
      }
    }
  }

  await saveWorkbook(workbook, filePath);
  const targetRange =
    range ??
    `${startCell}:${cellAddress(origin + grid.length - 1, originCol + grid[0].length - 1)}`;
  return operationSuccess(
    `Wrote ${grid.length} row(s) to sheet '${worksheet.name}' at ${targetRange}`
  );
}
