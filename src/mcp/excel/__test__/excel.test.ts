import { describe, expect, it, beforeAll, afterAll } from "@jest/globals";
import * as fs from "fs";
import * as path from "path";
import { fileURLToPath } from "url";
import {
  parseA1Range,
  createFile,
  readSheet,
  getMetadata,
  writeRange,
  formatRange,
  copySheet,
  renameSheet,
  setSheetVisibility,
  setDataValidation,
  setDimensions,
} from "../excel.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const tempFilePath = path.join(__dirname, "temp_test_workbook.xlsx");

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
    expect(parseA1Range("A1:C3")).toEqual({
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
