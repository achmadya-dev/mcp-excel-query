import {
  applyStyleToCell,
  loadWorkbook,
  operationSuccess,
  parseA1Range,
  resolveRangeBounds,
  resolveWorksheet,
  saveWorkbook,
} from "./utils.js";

export async function formatRange(params: {
  filePath: string;
  sheetName?: string;
  startCell: string;
  endCell?: string;
  range?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontSize?: number;
  fontColor?: string;
  bgColor?: string;
  borderStyle?: string;
  borderColor?: string;
  numberFormat?: string;
  alignment?: string;
  wrapText?: boolean;
  mergeCells?: boolean;
  styles?: (Record<string, unknown> | null)[][];
}) {
  const workbook = await loadWorkbook(params.filePath);
  const worksheet = resolveWorksheet(workbook, params.sheetName);

  const bounds = params.range
    ? parseA1Range(params.range)
    : resolveRangeBounds(params.startCell, params.endCell);

  const flatStyle: Record<string, unknown> = {};
  if (params.bold !== undefined || params.italic !== undefined || params.underline !== undefined || params.fontSize !== undefined || params.fontColor) {
    flatStyle.font = {
      ...(params.bold !== undefined ? { bold: params.bold } : {}),
      ...(params.italic !== undefined ? { italic: params.italic } : {}),
      ...(params.underline !== undefined ? { underline: params.underline } : {}),
      ...(params.fontSize !== undefined ? { size: params.fontSize } : {}),
      ...(params.fontColor ? { color: params.fontColor } : {}),
    };
  }
  if (params.bgColor) {
    flatStyle.fill = { fgColor: params.bgColor };
  }
  if (params.borderStyle) {
    const side = { style: params.borderStyle, ...(params.borderColor ? { color: params.borderColor } : {}) };
    flatStyle.border = { top: side, left: side, bottom: side, right: side };
  }
  if (params.alignment || params.wrapText !== undefined) {
    flatStyle.alignment = {
      ...(params.alignment ? { horizontal: params.alignment } : {}),
      ...(params.wrapText !== undefined ? { wrapText: params.wrapText } : {}),
    };
  }
  if (params.numberFormat) {
    flatStyle.numFmt = params.numberFormat;
  }

  const rowCount = bounds.endRow - bounds.startRow + 1;
  const colCount = bounds.endCol - bounds.startCol + 1;

  if (params.mergeCells) {
    worksheet.mergeCells(bounds.startRow, bounds.startCol, bounds.endRow, bounds.endCol);
  }

  for (let r = 0; r < rowCount; r++) {
    for (let c = 0; c < colCount; c++) {
      const cell = worksheet.getCell(bounds.startRow + r, bounds.startCol + c);
      let styleToApply: Record<string, unknown> | null = flatStyle;

      if (params.styles) {
        const gridStyle = params.styles[r]?.[c];
        if (gridStyle === null) continue;
        if (gridStyle) styleToApply = gridStyle;
      }

      if (styleToApply && Object.keys(styleToApply).length > 0) {
        applyStyleToCell(cell, styleToApply);
      }
    }
  }

  await saveWorkbook(workbook, params.filePath);
  return operationSuccess(
    `Formatted range ${bounds.startRow},${bounds.startCol}:${bounds.endRow},${bounds.endCol} on sheet '${worksheet.name}'`
  );
}
