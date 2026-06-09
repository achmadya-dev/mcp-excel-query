import ExcelJS from "exceljs";
import * as fs from "fs";
import {
  extractDateFromCellValue,
  formatDateValue,
  parseDateValue,
  resolveDateFormat,
} from "./date.js";

export type A1Bounds = {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
};

export type CellValue = string | number | boolean | null;

export function getPagingCellsLimit(): number {
  const raw = process.env.EXCEL_PAGING_CELLS_LIMIT;
  if (raw === undefined || raw === "") return 4000;
  const n = parseInt(raw, 10);
  return Number.isFinite(n) && n > 0 ? n : 4000;
}

export function assertFileExists(filePath: string): void {
  if (!fs.existsSync(filePath)) {
    throw new Error(`File not found: ${filePath}`);
  }
}

export async function loadWorkbook(filePath: string): Promise<ExcelJS.Workbook> {
  assertFileExists(filePath);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  return workbook;
}

export async function saveWorkbook(
  workbook: ExcelJS.Workbook,
  filePath: string
): Promise<void> {
  await workbook.xlsx.writeFile(filePath);
}

export function resolveWorksheet(
  workbook: ExcelJS.Workbook,
  sheetName?: string
): ExcelJS.Worksheet {
  const worksheet = sheetName
    ? workbook.getWorksheet(sheetName)
    : workbook.worksheets[0];
  if (!worksheet) {
    throw new Error(`Worksheet not found: "${sheetName ?? "Index 0"}"`);
  }
  return worksheet;
}

export function colLetterToNum(letter: string): number {
  let num = 0;
  for (let i = 0; i < letter.length; i++) {
    const charCode = letter.toUpperCase().charCodeAt(i);
    if (charCode < 65 || charCode > 90) {
      throw new Error(`Invalid column character in letter: ${letter}`);
    }
    num = num * 26 + (charCode - 64);
  }
  return num;
}

export function numToColLetter(col: number): string {
  let temp = "";
  let c = col;
  while (c > 0) {
    const rem = (c - 1) % 26;
    temp = String.fromCharCode(65 + rem) + temp;
    c = Math.floor((c - 1) / 26);
  }
  return temp || "A";
}

export function cellAddress(row: number, col: number): string {
  return `${numToColLetter(col)}${row}`;
}

export function boundsToRange(bounds: A1Bounds): string {
  const start = cellAddress(bounds.startRow, bounds.startCol);
  if (bounds.startRow === bounds.endRow && bounds.startCol === bounds.endCol) {
    return start;
  }
  return `${start}:${cellAddress(bounds.endRow, bounds.endCol)}`;
}

export function parseA1Range(rangeStr: string): A1Bounds {
  const trimmed = rangeStr.trim().replace(/\$/g, "");
  const match = trimmed.match(/^([A-Z]+)([0-9]+):([A-Z]+)([0-9]+)$/i);
  if (!match) {
    const singleMatch = trimmed.match(/^([A-Z]+)([0-9]+)$/i);
    if (singleMatch) {
      const col = colLetterToNum(singleMatch[1]);
      const row = parseInt(singleMatch[2], 10);
      return { startRow: row, startCol: col, endRow: row, endCol: col };
    }
    throw new Error(`Invalid A1 range or cell format: "${rangeStr}"`);
  }
  const startCol = colLetterToNum(match[1]);
  const startRow = parseInt(match[2], 10);
  const endCol = colLetterToNum(match[3]);
  const endRow = parseInt(match[4], 10);

  return {
    startRow: Math.min(startRow, endRow),
    startCol: Math.min(startCol, endCol),
    endRow: Math.max(startRow, endRow),
    endCol: Math.max(startCol, endCol),
  };
}

export function resolveRangeBounds(
  startCell: string,
  endCell?: string
): A1Bounds {
  if (endCell) {
    const start = parseA1Range(startCell);
    const end = parseA1Range(endCell);
    return {
      startRow: Math.min(start.startRow, end.startRow),
      startCol: Math.min(start.startCol, end.startCol),
      endRow: Math.max(start.endRow, end.endRow),
      endCol: Math.max(start.endCol, end.endCol),
    };
  }
  return parseA1Range(startCell);
}

export function getCleanCellValue(cellValue: unknown): CellValue | string {
  if (cellValue === null || cellValue === undefined) return null;
  if (cellValue instanceof Date) return cellValue.toISOString();

  if (typeof cellValue === "object") {
    if ("result" in cellValue) {
      return getCleanCellValue((cellValue as { result: unknown }).result);
    }
    if ("richText" in cellValue && Array.isArray((cellValue as { richText: unknown[] }).richText)) {
      return (cellValue as { richText: { text?: string }[] }).richText
        .map((t) => t.text || "")
        .join("");
    }
    if ("text" in cellValue) {
      return String((cellValue as { text: unknown }).text);
    }
    return JSON.stringify(cellValue) as string;
  }
  return cellValue as CellValue;
}

export function getCellFormula(cellValue: unknown): string | null {
  if (cellValue === null || cellValue === undefined) return null;
  if (typeof cellValue === "object" && "formula" in cellValue) {
    const formula = (cellValue as { formula: string }).formula;
    return formula.startsWith("=") ? formula : `=${formula}`;
  }
  if (typeof cellValue === "string" && cellValue.startsWith("=")) {
    return cellValue;
  }
  return null;
}

export function readCellDisplayValue(
  cellValue: unknown,
  showFormula: boolean,
  opts?: { dateFormat?: string; cellNumFmt?: string }
): CellValue | string {
  if (showFormula) {
    const formula = getCellFormula(cellValue);
    if (formula) return formula;
  }

  const resolvedFormat = resolveDateFormat(opts?.dateFormat, opts?.cellNumFmt);
  if (resolvedFormat) {
    const date = extractDateFromCellValue(cellValue);
    if (date) return formatDateValue(date, resolvedFormat);
  }

  return getCleanCellValue(cellValue);
}

export function extractCellStyle(cell: ExcelJS.Cell): Record<string, unknown> | null {
  const style: Record<string, unknown> = {};
  if (cell.font && Object.keys(cell.font).length > 0) style.font = { ...cell.font };
  if (cell.fill && Object.keys(cell.fill).length > 0) style.fill = { ...cell.fill };
  if (cell.alignment && Object.keys(cell.alignment).length > 0) {
    style.alignment = { ...cell.alignment };
  }
  if (cell.border && Object.keys(cell.border).length > 0) style.border = { ...cell.border };
  if (cell.numFmt) style.numFmt = cell.numFmt;
  return Object.keys(style).length > 0 ? style : null;
}

export function setCellValue(cell: ExcelJS.Cell, value: unknown): void {
  if (value === null || value === undefined) {
    cell.value = null;
    return;
  }
  if (typeof value === "string" && value.startsWith("=")) {
    cell.value = { formula: value.slice(1) };
    return;
  }

  const parsedDate = parseDateValue(value);
  if (parsedDate) {
    cell.value = parsedDate;
    return;
  }

  cell.value = value as ExcelJS.CellValue;
}

export function resolveColor(hex: string) {
  const clean = hex.trim().replace(/^#/, "");
  if (clean.length === 6) {
    return { argb: "FF" + clean.toUpperCase() };
  }
  if (clean.length === 8) {
    return { argb: clean.toUpperCase() };
  }
  return { argb: clean };
}

export function applyStyleToCell(excelCell: ExcelJS.Cell, style: Record<string, unknown>): void {
  if (!style) return;

  if (style.font && typeof style.font === "object") {
    const font = { ...(style.font as Record<string, unknown>) };
    if (typeof font.color === "string") {
      font.color = resolveColor(font.color);
    }
    excelCell.font = font as unknown as ExcelJS.Font;
  }

  if (style.fill && typeof style.fill === "object") {
    const fillSrc = style.fill as Record<string, unknown>;
    const fill: Record<string, unknown> = {};
    fill.type = fillSrc.type || "pattern";
    fill.pattern = fillSrc.pattern || "solid";
    if (typeof fillSrc.fgColor === "string") {
      fill.fgColor = resolveColor(fillSrc.fgColor);
    }
    if (typeof fillSrc.bgColor === "string") {
      fill.bgColor = resolveColor(fillSrc.bgColor);
    }
    excelCell.fill = fill as unknown as ExcelJS.Fill;
  }

  if (style.alignment && typeof style.alignment === "object") {
    excelCell.alignment = { ...(style.alignment as ExcelJS.Alignment) };
  }

  if (style.border && typeof style.border === "object") {
    const border: Partial<ExcelJS.Borders> = {};
    const sides = ["top", "left", "bottom", "right"] as const;
    const borderSrc = style.border as Record<string, { style?: string; color?: string } | undefined>;
    for (const side of sides) {
      if (borderSrc[side]) {
        border[side] = {
          style: borderSrc[side]!.style as ExcelJS.BorderStyle,
          ...(borderSrc[side]!.color
            ? { color: resolveColor(borderSrc[side]!.color!) }
            : {}),
        };
      }
    }
    excelCell.border = border as ExcelJS.Borders;
  }

  if (typeof style.numFmt === "string") {
    excelCell.numFmt = style.numFmt;
  }
}

export function operationSuccess(message: string) {
  return { success: true as const, message };
}
