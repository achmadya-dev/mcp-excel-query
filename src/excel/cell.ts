import type ExcelJS from "exceljs";
import {
  extractDateFromCellValue,
  formatDateValue,
  parseDateValue,
  resolveDateFormat,
} from "./date.js";

export type CellValue = string | number | boolean | null;

function toArgb(hex: string) {
  const clean = hex.trim().replace(/^#/, "");
  if (clean.length === 6) {
    return { argb: "FF" + clean.toUpperCase() };
  }
  if (clean.length === 8) {
    return { argb: clean.toUpperCase() };
  }
  return { argb: clean };
}

function getCleanValue(cellValue: unknown): CellValue | string {
  if (cellValue === null || cellValue === undefined) return null;
  if (cellValue instanceof Date) return cellValue.toISOString();

  if (typeof cellValue === "object") {
    if ("result" in cellValue) {
      return getCleanValue((cellValue as { result: unknown }).result);
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

function getFormula(cellValue: unknown): string | null {
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

export class Cell {
  static readDisplay(
    cellValue: unknown,
    showFormula: boolean,
    opts?: { dateFormat?: string; cellNumFmt?: string }
  ): CellValue | string {
    if (showFormula) {
      const formula = getFormula(cellValue);
      if (formula) return formula;
    }

    const resolvedFormat = resolveDateFormat(opts?.dateFormat, opts?.cellNumFmt);
    if (resolvedFormat) {
      const date = extractDateFromCellValue(cellValue);
      if (date) return formatDateValue(date, resolvedFormat);
    }

    return getCleanValue(cellValue);
  }

  static extractStyle(cell: ExcelJS.Cell): Record<string, unknown> | null {
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

  static set(cell: ExcelJS.Cell, value: unknown): void {
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

  static applyStyle(excelCell: ExcelJS.Cell, style: Record<string, unknown>): void {
    if (!style) return;

    if (style.font && typeof style.font === "object") {
      const font = { ...(style.font as Record<string, unknown>) };
      if (typeof font.color === "string") {
        font.color = toArgb(font.color);
      }
      excelCell.font = font as unknown as ExcelJS.Font;
    }

    if (style.fill && typeof style.fill === "object") {
      const fillSrc = style.fill as Record<string, unknown>;
      const fill: Record<string, unknown> = {};
      fill.type = fillSrc.type || "pattern";
      fill.pattern = fillSrc.pattern || "solid";
      if (typeof fillSrc.fgColor === "string") {
        fill.fgColor = toArgb(fillSrc.fgColor);
      }
      if (typeof fillSrc.bgColor === "string") {
        fill.bgColor = toArgb(fillSrc.bgColor);
      }
      excelCell.fill = fill as unknown as ExcelJS.Fill;
    }

    if (style.alignment && typeof style.alignment === "object") {
      excelCell.alignment = { ...(style.alignment as ExcelJS.Alignment) };
    }

    if (style.border && typeof style.border === "object") {
      const border: Partial<ExcelJS.Borders> = {};
      const sides = ["top", "left", "bottom", "right"] as const;
      const borderSrc = style.border as Record<
        string,
        { style?: string; color?: string } | undefined
      >;
      for (const side of sides) {
        if (borderSrc[side]) {
          border[side] = {
            style: borderSrc[side]!.style as ExcelJS.BorderStyle,
            ...(borderSrc[side]!.color ? { color: toArgb(borderSrc[side]!.color!) } : {}),
          };
        }
      }
      excelCell.border = border as ExcelJS.Borders;
    }

    if (typeof style.numFmt === "string") {
      excelCell.numFmt = style.numFmt;
    }
  }
}
