import {
  colLetterToNum,
  loadWorkbook,
  numToColLetter,
  operationSuccess,
  parseA1Range,
  resolveRangeBounds,
  resolveWorksheet,
  saveWorkbook,
} from "./utils.js";

function resolveColumnIndex(column: number | string): number {
  if (typeof column === "number") {
    if (!Number.isInteger(column) || column < 1) {
      throw new Error(`Invalid column index: ${column}`);
    }
    return column;
  }
  return colLetterToNum(column);
}

function collectTargetRows(params: {
  rows?: number[];
  range?: string;
  startCell?: string;
  endCell?: string;
}): number[] {
  if (params.rows && params.rows.length > 0) {
    return params.rows;
  }

  if (params.range) {
    const bounds = parseA1Range(params.range);
    return Array.from(
      { length: bounds.endRow - bounds.startRow + 1 },
      (_, i) => bounds.startRow + i
    );
  }

  if (params.startCell) {
    const bounds = resolveRangeBounds(params.startCell, params.endCell);
    return Array.from(
      { length: bounds.endRow - bounds.startRow + 1 },
      (_, i) => bounds.startRow + i
    );
  }

  return [];
}

function collectTargetColumns(params: {
  columns?: Array<number | string>;
  range?: string;
  startCell?: string;
  endCell?: string;
}): number[] {
  if (params.columns && params.columns.length > 0) {
    return params.columns.map(resolveColumnIndex);
  }

  if (params.range) {
    const bounds = parseA1Range(params.range);
    return Array.from(
      { length: bounds.endCol - bounds.startCol + 1 },
      (_, i) => bounds.startCol + i
    );
  }

  if (params.startCell) {
    const bounds = resolveRangeBounds(params.startCell, params.endCell);
    return Array.from(
      { length: bounds.endCol - bounds.startCol + 1 },
      (_, i) => bounds.startCol + i
    );
  }

  return [];
}

export async function setDimensions(params: {
  filePath: string;
  sheetName?: string;
  range?: string;
  startCell?: string;
  endCell?: string;
  rows?: number[];
  columns?: Array<number | string>;
  rowHeight?: number;
  columnWidth?: number;
  defaultRowHeight?: number;
  defaultColumnWidth?: number;
}) {
  const { filePath, sheetName, rowHeight, columnWidth, defaultRowHeight, defaultColumnWidth } =
    params;

  if (
    rowHeight === undefined &&
    columnWidth === undefined &&
    defaultRowHeight === undefined &&
    defaultColumnWidth === undefined
  ) {
    throw new Error(
      "Provide rowHeight, columnWidth, defaultRowHeight, and/or defaultColumnWidth"
    );
  }

  const workbook = await loadWorkbook(filePath);
  const worksheet = resolveWorksheet(workbook, sheetName);

  const targetRows = collectTargetRows(params);
  const targetColumns = collectTargetColumns(params);

  if (rowHeight !== undefined) {
    if (targetRows.length === 0) {
      throw new Error("rowHeight requires rows, range, or startCell");
    }
    for (const rowNumber of targetRows) {
      worksheet.getRow(rowNumber).height = rowHeight;
    }
  }

  if (columnWidth !== undefined) {
    if (targetColumns.length === 0) {
      throw new Error("columnWidth requires columns, range, or startCell");
    }
    for (const colNumber of targetColumns) {
      worksheet.getColumn(colNumber).width = columnWidth;
    }
  }

  if (defaultRowHeight !== undefined) {
    worksheet.properties.defaultRowHeight = defaultRowHeight;
  }

  if (defaultColumnWidth !== undefined) {
    worksheet.properties.defaultColWidth = defaultColumnWidth;
  }

  await saveWorkbook(workbook, filePath);

  const parts: string[] = [];
  if (rowHeight !== undefined) {
    parts.push(`row height ${rowHeight} on rows ${targetRows.join(", ")}`);
  }
  if (columnWidth !== undefined) {
    const labels = targetColumns.map((c) => numToColLetter(c));
    parts.push(`column width ${columnWidth} on columns ${labels.join(", ")}`);
  }
  if (defaultRowHeight !== undefined) {
    parts.push(`default row height ${defaultRowHeight}`);
  }
  if (defaultColumnWidth !== undefined) {
    parts.push(`default column width ${defaultColumnWidth}`);
  }

  return operationSuccess(`Updated ${parts.join("; ")} on sheet '${worksheet.name}'`);
}
