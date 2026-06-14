import { defineTool, ToolError } from "@achmadya-dev/mcp-core";
import { z } from "zod";
import type ExcelJS from "exceljs";
import { Excel } from "../excel/excel.js";

const filePathField = z.string().describe("Absolute path to the local Excel file (.xlsx)");

const copySheetInputSchema = z.object({
  filePath: filePathField,
  sourceSheet: z.string(),
  targetSheet: z.string(),
});

const operationResultOutputShape = z.object({
  success: z.boolean(),
  message: z.string(),
});

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

export const excel_copy_sheet = defineTool({
  name: "excel_copy_sheet",
  description: "Copy an existing worksheet to a new sheet name.",
  inputSchema: copySheetInputSchema,
  outputSchema: operationResultOutputShape,
  handler: async (args) => {
    try {
      const { filePath, sourceSheet, targetSheet } = args;
      const excel = await Excel.open(filePath);
      const source = excel.sheet(sourceSheet).raw;

      if (excel.hasSheet(targetSheet)) {
        throw new Error(`Target worksheet already exists: "${targetSheet}"`);
      }

      duplicateWorksheet(excel.workbook, source, targetSheet);
      await excel.save();
      return {
        success: true as const,
        message: `Sheet '${sourceSheet}' copied to '${targetSheet}'`,
      };
    } catch (error) {
      const msg = error instanceof Error ? error.message : String(error);
      throw new ToolError("Failed to copy sheet: " + msg);
    }
  },
});
