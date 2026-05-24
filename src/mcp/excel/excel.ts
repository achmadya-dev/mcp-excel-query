import ExcelJS from "exceljs";
import * as fs from "fs";
import * as path from "path";

// Helper to convert column letter to 1-based column index (e.g. "A" -> 1, "AA" -> 27)
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

// Helper to convert 1-based column index to letter (e.g. 1 -> "A", 27 -> "AA")
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

// Helper to parse A1 range notation (e.g., "A1:C10" or "B5")
export function parseA1Range(rangeStr: string) {
  const trimmed = rangeStr.trim().replace(/\$/g, ""); // strip absolute signs
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

// Helper to clean ExcelJS cell values to JSON-safe primitives
export function getCleanCellValue(cellValue: unknown): any {
  if (cellValue === null || cellValue === undefined) return null;
  if (cellValue instanceof Date) return cellValue.toISOString();
  
  if (typeof cellValue === "object") {
    // Formula Cell
    if ("result" in cellValue) {
      return getCleanCellValue(cellValue.result);
    }
    // Rich Text Cell
    if ("richText" in cellValue && Array.isArray(cellValue.richText)) {
      return cellValue.richText.map((t: any) => t.text || "").join("");
    }
    // Hyperlink Cell
    if ("text" in cellValue) {
      return cellValue.text;
    }
    // Safe fallback for other objects
    return JSON.stringify(cellValue);
  }
  return cellValue;
}

// 1. List worksheet names inside an Excel file
export async function listSheets(filePath: string) {
  if (!fs.existsSync(filePath)) {
    throw new Error(`File not found: ${filePath}`);
  }
  
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  
  return {
    sheets: workbook.worksheets.map((ws) => ws.name),
  };
}

// 2. Read sheet rows with optional range, headerRow, and pagination
export async function readSheet(params: {
  filePath: string;
  sheetName?: string;
  range?: string;
  headerRow?: number;
  limit?: number;
  offset?: number;
}) {
  const { filePath, sheetName, range, headerRow = 1, limit, offset } = params;

  if (!fs.existsSync(filePath)) {
    throw new Error(`File not found: ${filePath}`);
  }

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  let worksheet = sheetName ? workbook.getWorksheet(sheetName) : workbook.worksheets[0];
  if (!worksheet) {
    throw new Error(`Worksheet not found: "${sheetName || "Index 0"}"`);
  }

  // Determine bounds
  let startRow = 1;
  let endRow = worksheet.rowCount;
  let startCol = 1;
  let endCol = worksheet.actualColumnCount || 1;

  if (range) {
    const parsedRange = parseA1Range(range);
    startRow = parsedRange.startRow;
    endRow = parsedRange.endRow;
    startCol = parsedRange.startCol;
    endCol = parsedRange.endCol;
  }

  // Extract and clean headers
  const headers: string[] = [];
  const headerRowIdx = headerRow;

  if (headerRowIdx > 0 && headerRowIdx <= worksheet.rowCount) {
    const row = worksheet.getRow(headerRowIdx);
    for (let c = startCol; c <= endCol; c++) {
      const cellVal = getCleanCellValue(row.getCell(c).value);
      headers.push(
        cellVal !== null && cellVal !== undefined && cellVal !== ""
          ? String(cellVal)
          : numToColLetter(c)
      );
    }
  } else {
    // Generate generic header names A, B, C...
    for (let c = startCol; c <= endCol; c++) {
      headers.push(numToColLetter(c));
    }
  }

  // Extract data rows
  const allRows: Record<string, any>[] = [];
  const startDataRow = headerRowIdx > 0 ? Math.max(startRow, headerRowIdx + 1) : startRow;

  for (let r = startDataRow; r <= endRow; r++) {
    const row = worksheet.getRow(r);
    const rowData: Record<string, any> = {};
    let hasValue = false;

    for (let c = startCol; c <= endCol; c++) {
      const cellVal = getCleanCellValue(row.getCell(c).value);
      const headerKey = headers[c - startCol];
      rowData[headerKey] = cellVal;
      if (cellVal !== null && cellVal !== undefined && cellVal !== "") {
        hasValue = true;
      }
    }

    // Save if row has any values or we are scanning an explicit range
    if (hasValue || range) {
      allRows.push(rowData);
    }
  }

  // Apply limit and offset pagination
  const totalRows = allRows.length;
  const off = offset ?? 0;
  const lim = limit ?? totalRows;
  const paginatedRows = allRows.slice(off, off + lim);

  return {
    sheetName: worksheet.name,
    headers: headerRowIdx > 0 ? headers : null,
    rows: paginatedRows,
    totalRows,
  };
}

// 3. Read metadata of an Excel file
export async function getMetadata(filePath: string) {
  if (!fs.existsSync(filePath)) {
    throw new Error(`File not found: ${filePath}`);
  }

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  const sheets = workbook.worksheets.map((ws) => ({
    name: ws.name,
    rowCount: ws.rowCount,
    columnCount: ws.actualColumnCount || 0,
  }));

  return {
    filePath,
    creator: workbook.creator || null,
    lastModifiedBy: workbook.lastModifiedBy || null,
    created: workbook.created ? workbook.created.toISOString() : null,
    modified: workbook.modified ? workbook.modified.toISOString() : null,
    sheets,
  };
}

// 4. Create a new Excel file
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

  await workbook.xlsx.writeFile(filePath);

  return {
    success: true,
    message: `Excel file created successfully at ${filePath}`,
  };
}

// 5. Create a new worksheet in an existing file
export async function createSheet(params: {
  filePath: string;
  sheetName: string;
}) {
  const { filePath, sheetName } = params;

  if (!fs.existsSync(filePath)) {
    throw new Error(`File not found: ${filePath}`);
  }

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  if (workbook.getWorksheet(sheetName)) {
    throw new Error(`Worksheet already exists: "${sheetName}"`);
  }

  workbook.addWorksheet(sheetName);
  await workbook.xlsx.writeFile(filePath);

  return {
    success: true,
    message: `Sheet '${sheetName}' created successfully in ${filePath}`,
  };
}

// Helper to resolve color hex string to ExcelJS ARGB color object
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

// Helper to apply custom style object to an ExcelJS cell
export function applyStyleToCell(excelCell: any, style: any) {
  if (!style) return;

  if (style.font) {
    const font: any = { ...style.font };
    if (style.font.color) {
      font.color = resolveColor(style.font.color);
    }
    excelCell.font = font;
  }

  if (style.fill) {
    const fill: any = {};
    fill.type = style.fill.type || "pattern";
    fill.pattern = style.fill.pattern || "solid";
    if (style.fill.fgColor) {
      fill.fgColor = resolveColor(style.fill.fgColor);
    }
    if (style.fill.bgColor) {
      fill.bgColor = resolveColor(style.fill.bgColor);
    }
    excelCell.fill = fill;
  }

  if (style.alignment) {
    excelCell.alignment = { ...style.alignment };
  }

  if (style.border) {
    const border: any = {};
    const sides = ["top", "left", "bottom", "right"] as const;
    for (const side of sides) {
      if (style.border[side]) {
        border[side] = {
          style: style.border[side].style,
          ...(style.border[side].color ? { color: resolveColor(style.border[side].color) } : {})
        };
      }
    }
    excelCell.border = border;
  }
}

// 6. Write value and/or apply style to a specific cell
export async function writeCell(params: {
  filePath: string;
  sheetName?: string;
  cell: string;
  value?: any;
  style?: any;
}) {
  const { filePath, sheetName, cell, value, style } = params;

  if (!fs.existsSync(filePath)) {
    throw new Error(`File not found: ${filePath}`);
  }

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  let worksheet = sheetName ? workbook.getWorksheet(sheetName) : workbook.worksheets[0];
  if (!worksheet) {
    throw new Error(`Worksheet not found: "${sheetName || "Index 0"}"`);
  }

  const excelCell = worksheet.getCell(cell);
  
  if (value !== undefined) {
    excelCell.value = value;
  }

  if (style) {
    applyStyleToCell(excelCell, style);
  }

  await workbook.xlsx.writeFile(filePath);

  return {
    success: true,
    message: `Cell ${cell} in sheet '${worksheet.name}' successfully updated (value written: ${value !== undefined}, style applied: ${!!style})`,
  };
}

// 7. Append multiple rows of data
export async function appendRows(params: {
  filePath: string;
  sheetName?: string;
  rows: any[][];
}) {
  const { filePath, sheetName, rows } = params;

  if (!fs.existsSync(filePath)) {
    throw new Error(`File not found: ${filePath}`);
  }

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  let worksheet = sheetName ? workbook.getWorksheet(sheetName) : workbook.worksheets[0];
  if (!worksheet) {
    throw new Error(`Worksheet not found: "${sheetName || "Index 0"}"`);
  }

  for (const rowData of rows) {
    worksheet.addRow(rowData);
  }

  await workbook.xlsx.writeFile(filePath);

  return {
    success: true,
    message: `Successfully appended ${rows.length} row(s) to sheet '${worksheet.name}'`,
  };
}
