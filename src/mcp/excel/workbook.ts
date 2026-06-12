import ExcelJS from "exceljs";
import * as fs from "fs";
import * as path from "path";
import {
  boundsToRange,
  loadWorkbook,
  operationSuccess,
  saveWorkbook,
} from "./utils.js";

export async function getMetadata(params: {
  filePath: string;
  includeRanges?: boolean;
}) {
  const { filePath, includeRanges = false } = params;
  const workbook = await loadWorkbook(filePath);

  const sheets = workbook.worksheets.map((ws) => {
    const base = {
      name: ws.name,
      rowCount: ws.rowCount,
      columnCount: ws.actualColumnCount || 0,
      state: ws.state ?? "visible",
    };
    if (!includeRanges) return base;

    let usedRange: string | null = null;
    if (ws.rowCount > 0 && (ws.actualColumnCount || 0) > 0) {
      usedRange = boundsToRange({
        startRow: 1,
        startCol: 1,
        endRow: ws.rowCount,
        endCol: ws.actualColumnCount || 1,
      });
    }

    const mergedCells = (ws.model.merges ?? []).map((m) => String(m));
    return { ...base, usedRange, mergedCells };
  });

  return {
    filePath,
    creator: workbook.creator || null,
    lastModifiedBy: workbook.lastModifiedBy || null,
    created: workbook.created ? workbook.created.toISOString() : null,
    modified: workbook.modified ? workbook.modified.toISOString() : null,
    sheets,
  };
}

export async function createFile(params: {
  filePath: string;
  sheetName?: string;
  headers?: string[];
}) {
  const { filePath, sheetName = "Sheet1", headers } = params;

  const dir = path.dirname(filePath);
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }

  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet(sheetName);

  if (headers && headers.length > 0) {
    worksheet.addRow(headers);
  }

  await saveWorkbook(workbook, filePath);
  return operationSuccess(`Excel file created successfully at ${filePath}`);
}
