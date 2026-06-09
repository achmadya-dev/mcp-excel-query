import ExcelJS from "exceljs";
import {
  loadWorkbook,
  operationSuccess,
  resolveWorksheet,
  saveWorkbook,
} from "./utils.js";

function duplicateWorksheet(
  workbook: ExcelJS.Workbook,
  source: ExcelJS.Worksheet,
  name: string
): ExcelJS.Worksheet {
  const newSheet = workbook.addWorksheet(name);
  newSheet.model = Object.assign({}, source.model, { name });

  source.eachRow((row, rowNumber) => {
    const newRow = newSheet.getRow(rowNumber);
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      const newCell = newRow.getCell(colNumber);
      newCell.value = cell.value;
      newCell.style = Object.assign({}, cell.style);
    });
  });

  for (const merge of source.model.merges ?? []) {
    newSheet.mergeCells(String(merge));
  }

  return newSheet;
}

export async function copySheet(params: {
  filePath: string;
  sourceSheet: string;
  targetSheet: string;
}) {
  const { filePath, sourceSheet, targetSheet } = params;
  const workbook = await loadWorkbook(filePath);
  const source = resolveWorksheet(workbook, sourceSheet);

  if (workbook.getWorksheet(targetSheet)) {
    throw new Error(`Target worksheet already exists: "${targetSheet}"`);
  }

  duplicateWorksheet(workbook, source, targetSheet);
  await saveWorkbook(workbook, filePath);
  return operationSuccess(`Sheet '${sourceSheet}' copied to '${targetSheet}'`);
}

export async function renameSheet(params: {
  filePath: string;
  oldName: string;
  newName: string;
}) {
  const { filePath, oldName, newName } = params;
  const workbook = await loadWorkbook(filePath);
  const worksheet = resolveWorksheet(workbook, oldName);

  if (workbook.getWorksheet(newName)) {
    throw new Error(`Worksheet already exists: "${newName}"`);
  }

  worksheet.name = newName;
  await saveWorkbook(workbook, filePath);
  return operationSuccess(`Sheet renamed from '${oldName}' to '${newName}'`);
}

export async function deleteSheet(params: {
  filePath: string;
  sheetName: string;
}) {
  const { filePath, sheetName } = params;
  const workbook = await loadWorkbook(filePath);

  if (workbook.worksheets.length <= 1) {
    throw new Error("Cannot delete the only worksheet in the workbook");
  }

  const worksheet = resolveWorksheet(workbook, sheetName);
  workbook.removeWorksheet(worksheet.id);
  await saveWorkbook(workbook, filePath);
  return operationSuccess(`Sheet '${sheetName}' deleted`);
}
