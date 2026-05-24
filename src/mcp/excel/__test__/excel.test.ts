import { describe, expect, it, beforeAll, afterAll } from "@jest/globals";
import * as fs from "fs";
import * as path from "path";
import { fileURLToPath } from "url";
import {
  colLetterToNum,
  numToColLetter,
  parseA1Range,
  createFile,
  createSheet,
  writeCell,
  appendRows,
  readSheet,
  getMetadata,
  listSheets,
} from "../excel.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const tempFilePath = path.join(__dirname, "temp_test_workbook.xlsx");

describe("A1 Notation Utilities", () => {
  it("converts column letters to numbers correctly", () => {
    expect(colLetterToNum("A")).toBe(1);
    expect(colLetterToNum("B")).toBe(2);
    expect(colLetterToNum("Z")).toBe(26);
    expect(colLetterToNum("AA")).toBe(27);
    expect(colLetterToNum("AB")).toBe(28);
  });

  it("converts column numbers to letters correctly", () => {
    expect(numToColLetter(1)).toBe("A");
    expect(numToColLetter(2)).toBe("B");
    expect(numToColLetter(26)).toBe("Z");
    expect(numToColLetter(27)).toBe("AA");
    expect(numToColLetter(28)).toBe("AB");
  });

  it("parses A1 notation range correctly", () => {
    expect(parseA1Range("A1")).toEqual({
      startRow: 1,
      startCol: 1,
      endRow: 1,
      endCol: 1,
    });
    expect(parseA1Range("A1:C10")).toEqual({
      startRow: 1,
      startCol: 1,
      endRow: 10,
      endCol: 3,
    });
    expect(parseA1Range("C10:A1")).toEqual({
      startRow: 1,
      startCol: 1,
      endRow: 10,
      endCol: 3,
    });
  });

  it("throws error for invalid A1 format", () => {
    expect(() => parseA1Range("Invalid")).toThrow();
  });
});

describe("Excel Read/Write Operations", () => {
  beforeAll(async () => {
    // Ensure clean state: delete temp file if exists
    if (fs.existsSync(tempFilePath)) {
      fs.unlinkSync(tempFilePath);
    }
  });

  afterAll(() => {
    // Cleanup temp file
    if (fs.existsSync(tempFilePath)) {
      fs.unlinkSync(tempFilePath);
    }
  });

  it("creates a new Excel file with headers", async () => {
    const res = await createFile({
      filePath: tempFilePath,
      sheetName: "Data",
      headers: ["ID", "Name", "Age"],
    });

    expect(res.success).toBe(true);
    expect(fs.existsSync(tempFilePath)).toBe(true);
  });

  it("lists sheets in the workbook", async () => {
    const res = await listSheets(tempFilePath);
    expect(res.sheets).toEqual(["Data"]);
  });

  it("appends rows of data to the worksheet", async () => {
    const res = await appendRows({
      filePath: tempFilePath,
      sheetName: "Data",
      rows: [
        [1, "Achmadya", 25],
        [2, "John Doe", 30],
      ],
    });

    expect(res.success).toBe(true);
  });

  it("reads back worksheet data correctly", async () => {
    const res = await readSheet({
      filePath: tempFilePath,
      sheetName: "Data",
      headerRow: 1,
    });

    expect(res.sheetName).toBe("Data");
    expect(res.headers).toEqual(["ID", "Name", "Age"]);
    expect(res.totalRows).toBe(2);
    expect(res.rows).toEqual([
      { ID: 1, Name: "Achmadya", Age: 25 },
      { ID: 2, Name: "John Doe", Age: 30 },
    ]);
  });

  it("reads with range bounds correctly", async () => {
    const res = await readSheet({
      filePath: tempFilePath,
      sheetName: "Data",
      range: "B1:C3",
      headerRow: 1,
    });

    expect(res.headers).toEqual(["Name", "Age"]);
    expect(res.rows).toEqual([
      { Name: "Achmadya", Age: 25 },
      { Name: "John Doe", Age: 30 },
    ]);
  });

  it("supports limit and offset pagination", async () => {
    const res = await readSheet({
      filePath: tempFilePath,
      sheetName: "Data",
      headerRow: 1,
      limit: 1,
      offset: 1,
    });

    expect(res.totalRows).toBe(2);
    expect(res.rows.length).toBe(1);
    expect(res.rows[0]).toEqual({ ID: 2, Name: "John Doe", Age: 30 });
  });

  it("writes to a specific cell and reads it back", async () => {
    const writeRes = await writeCell({
      filePath: tempFilePath,
      sheetName: "Data",
      cell: "B3",
      value: "Jane Smith",
    });
    expect(writeRes.success).toBe(true);

    const readRes = await readSheet({
      filePath: tempFilePath,
      sheetName: "Data",
      headerRow: 1,
    });
    expect(readRes.rows[1].Name).toBe("Jane Smith");
  });

  it("writes to a cell with custom style and verifies it", async () => {
    const res = await writeCell({
      filePath: tempFilePath,
      sheetName: "Data",
      cell: "A3",
      value: 99,
      style: {
        font: {
          bold: true,
          color: "FF0000",
        },
        fill: {
          fgColor: "FFFF00",
        },
      },
    });
    expect(res.success).toBe(true);

    const ExcelJS = (await import("exceljs")).default;
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(tempFilePath);
    const worksheet = workbook.getWorksheet("Data")!;
    const cell = worksheet.getCell("A3");
    
    expect(cell.value).toBe(99);
    expect(cell.font?.bold).toBe(true);
    expect(cell.font?.color?.argb).toBe("FFFF0000");
    expect(cell.fill?.type).toBe("pattern");
    expect((cell.fill as any).fgColor?.argb).toBe("FFFFFF00");
  });

  it("creates a new worksheet in the file", async () => {
    const res = await createSheet({
      filePath: tempFilePath,
      sheetName: "Summary",
    });
    expect(res.success).toBe(true);

    const sheetList = await listSheets(tempFilePath);
    expect(sheetList.sheets).toContain("Summary");
  });

  it("reads workbook metadata", async () => {
    const res = await getMetadata(tempFilePath);
    expect(res.filePath).toBe(tempFilePath);
    expect(res.sheets.length).toBe(2);
    expect(res.sheets.map((s) => s.name)).toContain("Data");
    expect(res.sheets.map((s) => s.name)).toContain("Summary");
  });
});
