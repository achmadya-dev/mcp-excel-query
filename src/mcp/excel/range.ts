import {
  loadWorkbook,
  operationSuccess,
  parseA1Range,
  resolveRangeBounds,
  resolveWorksheet,
  saveWorkbook,
} from "./utils.js";

export async function copyRange(params: {
  filePath: string;
  sheetName?: string;
  sourceStart: string;
  sourceEnd: string;
  targetStart: string;
  targetSheet?: string;
}) {
  const workbook = await loadWorkbook(params.filePath);
  const sourceWs = resolveWorksheet(workbook, params.sheetName);
  const targetWs = params.targetSheet
    ? resolveWorksheet(workbook, params.targetSheet)
    : sourceWs;

  const src = resolveRangeBounds(params.sourceStart, params.sourceEnd);
  const tgt = parseA1Range(params.targetStart);
  const rowOffset = tgt.startRow - src.startRow;
  const colOffset = tgt.startCol - src.startCol;

  for (let r = src.startRow; r <= src.endRow; r++) {
    for (let c = src.startCol; c <= src.endCol; c++) {
      const srcCell = sourceWs.getCell(r, c);
      const dstCell = targetWs.getCell(r + rowOffset, c + colOffset);
      dstCell.value = srcCell.value;
      dstCell.style = Object.assign({}, srcCell.style);
    }
  }

  await saveWorkbook(workbook, params.filePath);
  return operationSuccess(
    `Copied ${params.sourceStart}:${params.sourceEnd} to ${params.targetStart} on sheet '${targetWs.name}'`
  );
}

export async function deleteRange(params: {
  filePath: string;
  sheetName?: string;
  startCell: string;
  endCell: string;
  shiftDirection?: "up" | "left";
}) {
  const workbook = await loadWorkbook(params.filePath);
  const worksheet = resolveWorksheet(workbook, params.sheetName);
  const bounds = resolveRangeBounds(params.startCell, params.endCell);
  const shift = params.shiftDirection ?? "up";

  if (shift === "up") {
    worksheet.spliceRows(
      bounds.startRow,
      bounds.endRow - bounds.startRow + 1
    );
  } else {
    for (let r = bounds.startRow; r <= bounds.endRow; r++) {
      const row = worksheet.getRow(r);
      row.splice(bounds.startCol, bounds.endCol - bounds.startCol + 1);
    }
  }

  await saveWorkbook(workbook, params.filePath);
  return operationSuccess(
    `Deleted range ${params.startCell}:${params.endCell} (shift ${shift})`
  );
}

export async function unmergeCells(params: {
  filePath: string;
  sheetName?: string;
  startCell: string;
  endCell: string;
}) {
  const workbook = await loadWorkbook(params.filePath);
  const worksheet = resolveWorksheet(workbook, params.sheetName);
  const bounds = resolveRangeBounds(params.startCell, params.endCell);
  worksheet.unMergeCells(bounds.startRow, bounds.startCol, bounds.endRow, bounds.endCol);
  await saveWorkbook(workbook, params.filePath);
  return operationSuccess(`Unmerged cells ${params.startCell}:${params.endCell}`);
}

export async function insertRows(params: {
  filePath: string;
  sheetName?: string;
  startRow: number;
  count?: number;
}) {
  const { filePath, sheetName, startRow, count = 1 } = params;
  const workbook = await loadWorkbook(filePath);
  const worksheet = resolveWorksheet(workbook, sheetName);
  worksheet.spliceRows(startRow, 0, ...Array.from({ length: count }, () => []));
  await saveWorkbook(workbook, filePath);
  return operationSuccess(`Inserted ${count} row(s) at row ${startRow}`);
}

export async function insertColumns(params: {
  filePath: string;
  sheetName?: string;
  startCol: number;
  count?: number;
}) {
  const { filePath, sheetName, startCol, count = 1 } = params;
  const workbook = await loadWorkbook(filePath);
  const worksheet = resolveWorksheet(workbook, sheetName);

  worksheet.eachRow((row) => {
    row.splice(startCol, 0, ...Array(count).fill(null));
  });

  await saveWorkbook(workbook, filePath);
  return operationSuccess(`Inserted ${count} column(s) at column ${startCol}`);
}
