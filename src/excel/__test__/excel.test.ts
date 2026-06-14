import { describe, expect, it, beforeAll, afterAll } from "@jest/globals";
import type { RegisterableTool } from "@achmadya-dev/mcp-core";
import * as fs from "fs";
import * as path from "path";
import { fileURLToPath } from "url";
import { Range } from "../range.js";
import { excel_copy_sheet } from "../../tools/excel_copy_sheet.js";
import { excel_create_file } from "../../tools/excel_create_file.js";
import { excel_format_range } from "../../tools/excel_format_range.js";
import { excel_get_metadata } from "../../tools/excel_get_metadata.js";
import { excel_read_sheet } from "../../tools/excel_read_sheet.js";
import { excel_rename_sheet } from "../../tools/excel_rename_sheet.js";
import { excel_set_data_validation } from "../../tools/excel_set_data_validation.js";
import { excel_set_dimensions } from "../../tools/excel_set_dimensions.js";
import { excel_set_sheet_visibility } from "../../tools/excel_set_sheet_visibility.js";
import { excel_write_range } from "../../tools/excel_write_range.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const tempFilePath = path.join(__dirname, "temp_test_workbook.xlsx");

type OperationResult = { success: boolean; message: string };

type MetadataResult = {
  filePath: string;
  creator: string | null;
  lastModifiedBy: string | null;
  created: string | null;
  modified: string | null;
  sheets: Array<{
    name: string;
    rowCount: number;
    columnCount: number;
    state?: "visible" | "hidden" | "veryHidden";
    usedRange?: string | null;
    mergedCells?: string[];
  }>;
};

type ReadSheetResult = {
  sheetName: string;
  headers: string[] | null;
  rows: Record<string, unknown>[];
  totalRows: number;
  mergedCells?: string[];
  previewOnly?: boolean;
};

function typedHandler<TOut>(tool: RegisterableTool): (args: Record<string, unknown>) => Promise<TOut> {
  return tool.handler as (args: Record<string, unknown>) => Promise<TOut>;
}

const createFile = typedHandler<OperationResult>(excel_create_file);
const getMetadata = typedHandler<MetadataResult>(excel_get_metadata);
const writeRange = typedHandler<OperationResult>(excel_write_range);
const readSheet = typedHandler<ReadSheetResult>(excel_read_sheet);
const formatRange = typedHandler<OperationResult>(excel_format_range);
const copySheet = typedHandler<OperationResult>(excel_copy_sheet);
const renameSheet = typedHandler<OperationResult>(excel_rename_sheet);
const setSheetVisibility = typedHandler<OperationResult>(excel_set_sheet_visibility);
const setDataValidation = typedHandler<OperationResult>(excel_set_data_validation);
const setDimensions = typedHandler<OperationResult>(excel_set_dimensions);

describe("Excel operations", () => {
  beforeAll(() => {
    if (fs.existsSync(tempFilePath)) fs.unlinkSync(tempFilePath);
  });

  afterAll(() => {
    if (fs.existsSync(tempFilePath)) fs.unlinkSync(tempFilePath);
  });

  it("creates file and lists sheets via metadata", async () => {
    await createFile({
      filePath: tempFilePath,
      sheetName: "Data",
      headers: ["ID", "Name", "Age"],
    });
    const meta = await getMetadata({ filePath: tempFilePath });
    expect(meta.sheets.map((s) => s.name)).toEqual(["Data"]);
  });

  it("appends rows via writeRange append mode", async () => {
    const res = await writeRange({
      filePath: tempFilePath,
      sheetName: "Data",
      append: true,
      values: [
        [1, "Achmadya", 25],
        [2, "John Doe", 30],
      ],
    });
    expect(res.success).toBe(true);
  });

  it("reads worksheet data", async () => {
    const res = await readSheet({
      filePath: tempFilePath,
      sheetName: "Data",
      headerRow: 1,
    });
    expect(res.totalRows).toBe(2);
    expect(res.rows[0]).toEqual({ ID: 1, Name: "Achmadya", Age: 25 });
  });

  it("writes single cell formula via writeRange", async () => {
    await writeRange({
      filePath: tempFilePath,
      sheetName: "Data",
      startCell: "D2",
      values: [["=B2"]],
    });
    const read = await readSheet({
      filePath: tempFilePath,
      sheetName: "Data",
      range: "D2",
      headerRow: 0,
      showFormula: true,
    });
    expect(read.rows[0].D).toBe("=B2");
  });

  it("writes object data and formats range with merge", async () => {
    await writeRange({
      filePath: tempFilePath,
      sheetName: "Data",
      startCell: "E1",
      data: [{ City: "Jakarta" }],
    });
    const fmt = await formatRange({
      filePath: tempFilePath,
      sheetName: "Data",
      startCell: "F1",
      endCell: "G1",
      mergeCells: true,
      bold: true,
    });
    expect(fmt.success).toBe(true);

    const read = await readSheet({
      filePath: tempFilePath,
      sheetName: "Data",
      headerRow: 1,
    });
    expect(read.mergedCells!.length).toBeGreaterThan(0);
  });

  it("creates blank sheet via writeRange newSheet", async () => {
    const res = await writeRange({
      filePath: tempFilePath,
      newSheet: true,
      sheetName: "Summary",
    });
    expect(res.success).toBe(true);
    const meta = await getMetadata({ filePath: tempFilePath });
    expect(meta.sheets.map((s) => s.name)).toContain("Summary");
  });

  it("copies and renames sheets", async () => {
    await copySheet({
      filePath: tempFilePath,
      sourceSheet: "Data",
      targetSheet: "DataCopy",
    });
    await renameSheet({
      filePath: tempFilePath,
      oldName: "DataCopy",
      newName: "Archive",
    });
    const meta = await getMetadata({ filePath: tempFilePath });
    expect(meta.sheets.map((s) => s.name)).toContain("Archive");
  });

  it("parses A1 ranges", () => {
    expect(Range.parse("A1:C3")).toEqual({
      startRow: 1,
      startCol: 1,
      endRow: 3,
      endCol: 3,
    });
  });

  it("sets sheet visibility, data validation, and dimensions", async () => {
    await setSheetVisibility({
      filePath: tempFilePath,
      sheetName: "Archive",
      state: "hidden",
    });

    const meta = await getMetadata({ filePath: tempFilePath });
    const archive = meta.sheets.find((s) => s.name === "Archive");
    expect(archive?.state).toBe("hidden");

    await setDataValidation({
      filePath: tempFilePath,
      sheetName: "Data",
      startCell: "E2",
      endCell: "E3",
      type: "list",
      formulae: ['"A,B,C"'],
      allowBlank: true,
    });

    const withValidation = await readSheet({
      filePath: tempFilePath,
      sheetName: "Data",
      range: "E2",
      headerRow: 0,
    });
    expect(withValidation.rows[0].__validation_E).toBeDefined();

    await setDataValidation({
      filePath: tempFilePath,
      sheetName: "Data",
      startCell: "E2",
      endCell: "E3",
      clear: true,
    });

    const cleared = await readSheet({
      filePath: tempFilePath,
      sheetName: "Data",
      range: "E2",
      headerRow: 0,
    });
    expect(cleared.rows[0].__validation_E).toBeUndefined();

    const dim = await setDimensions({
      filePath: tempFilePath,
      sheetName: "Data",
      range: "A1:B2",
      rowHeight: 24,
      columnWidth: 16,
    });
    expect(dim.success).toBe(true);
  });

  it("writes human-readable dates and reads them with dateFormat", async () => {
    await writeRange({
      filePath: tempFilePath,
      sheetName: "Data",
      startCell: "H1",
      values: [["Date"], ["09 Jun 2026"]],
    });
    await formatRange({
      filePath: tempFilePath,
      sheetName: "Data",
      startCell: "H2",
      endCell: "H2",
      numberFormat: 'dd" "mmm" "yy',
    });

    const isoRead = await readSheet({
      filePath: tempFilePath,
      sheetName: "Data",
      range: "H2",
      headerRow: 0,
    });
    expect(String(isoRead.rows[0].H)).toMatch(/^2026-06-0[89]T\d{2}:\d{2}:\d{2}/);

    const formattedRead = await readSheet({
      filePath: tempFilePath,
      sheetName: "Data",
      range: "H2",
      headerRow: 0,
      dateFormat: "cell",
    });
    expect(formattedRead.rows[0].H).toBe("09 Jun 26");

    const customRead = await readSheet({
      filePath: tempFilePath,
      sheetName: "Data",
      range: "H2",
      headerRow: 0,
      dateFormat: "dd mmm yyyy",
    });
    expect(customRead.rows[0].H).toBe("09 Jun 2026");
  });
});
