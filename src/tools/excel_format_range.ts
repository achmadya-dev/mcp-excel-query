import { defineTool, ToolError } from "@achmadya-dev/mcp-core";
import { z } from "zod";
import { Cell } from "../excel/cell.js";
import { Excel } from "../excel/excel.js";

const filePathField = z.string().describe("Absolute path to the local Excel file (.xlsx)");

const sheetNameOptional = z
  .string()
  .optional()
  .describe("Name of the worksheet (defaults to the first sheet)");

const borderSideSchema = z
  .object({
    style: z.string().optional(),
    color: z.string().optional(),
  })
  .optional();
const cellStyleSchema = z.object({
  font: z
    .object({
      name: z.string().optional(),
      size: z.number().optional(),
      bold: z.boolean().optional(),
      italic: z.boolean().optional(),
      underline: z.boolean().optional(),
      color: z.string().optional(),
    })
    .optional(),
  fill: z
    .object({
      type: z.enum(["pattern", "gradient"]).optional(),
      pattern: z.string().optional(),
      fgColor: z.string().optional(),
      bgColor: z.string().optional(),
    })
    .optional(),
  alignment: z
    .object({
      horizontal: z.enum(["left", "center", "right", "justify", "distributed"]).optional(),
      vertical: z.enum(["top", "middle", "bottom", "justify", "distributed"]).optional(),
      wrapText: z.boolean().optional(),
    })
    .optional(),
  border: z
    .object({
      top: borderSideSchema,
      left: borderSideSchema,
      bottom: borderSideSchema,
      right: borderSideSchema,
    })
    .optional(),
  numFmt: z.string().optional(),
});

const formatRangeInputSchema = z.object({
  filePath: filePathField,
  sheetName: sheetNameOptional,
  startCell: z.string(),
  endCell: z.string().optional(),
  range: z.string().optional(),
  bold: z.boolean().optional(),
  italic: z.boolean().optional(),
  underline: z.boolean().optional(),
  fontSize: z.number().optional(),
  fontColor: z.string().optional(),
  bgColor: z.string().optional(),
  borderStyle: z.string().optional(),
  borderColor: z.string().optional(),
  numberFormat: z.string().optional(),
  alignment: z.string().optional(),
  wrapText: z.boolean().optional(),
  mergeCells: z.boolean().optional().describe("Merge the range while applying format"),
  styles: z.array(z.array(cellStyleSchema.nullable())).optional(),
});

const operationResultOutputShape = z.object({
  success: z.boolean(),
  message: z.string(),
});

export const excel_format_range = defineTool({
  name: "excel_format_range",
  description:
    "Format a cell range (font, fill, border, alignment, numFmt). Set mergeCells:true to merge. Use for styling after excel_write_range.",
  inputSchema: formatRangeInputSchema,
  outputSchema: operationResultOutputShape,
  handler: async (args) => {
    try {
      const excel = await Excel.open(args.filePath);
      const sheet = excel.sheet(args.sheetName);

      const bounds = args.range
        ? sheet.range(args.range)
        : args.endCell
          ? sheet.range(args.startCell, args.endCell)
          : sheet.range(args.startCell);

      const worksheet = sheet.raw;
      const flatStyle: Record<string, unknown> = {};
      if (
        args.bold !== undefined ||
        args.italic !== undefined ||
        args.underline !== undefined ||
        args.fontSize !== undefined ||
        args.fontColor
      ) {
        flatStyle.font = {
          ...(args.bold !== undefined ? { bold: args.bold } : {}),
          ...(args.italic !== undefined ? { italic: args.italic } : {}),
          ...(args.underline !== undefined ? { underline: args.underline } : {}),
          ...(args.fontSize !== undefined ? { size: args.fontSize } : {}),
          ...(args.fontColor ? { color: args.fontColor } : {}),
        };
      }
      if (args.bgColor) {
        flatStyle.fill = { fgColor: args.bgColor };
      }
      if (args.borderStyle) {
        const side = {
          style: args.borderStyle,
          ...(args.borderColor ? { color: args.borderColor } : {}),
        };
        flatStyle.border = { top: side, left: side, bottom: side, right: side };
      }
      if (args.alignment || args.wrapText !== undefined) {
        flatStyle.alignment = {
          ...(args.alignment ? { horizontal: args.alignment } : {}),
          ...(args.wrapText !== undefined ? { wrapText: args.wrapText } : {}),
        };
      }
      if (args.numberFormat) {
        flatStyle.numFmt = args.numberFormat;
      }

      const rowCount = bounds.endRow - bounds.startRow + 1;
      const colCount = bounds.endCol - bounds.startCol + 1;

      if (args.mergeCells) {
        worksheet.mergeCells(bounds.startRow, bounds.startCol, bounds.endRow, bounds.endCol);
      }

      for (let r = 0; r < rowCount; r++) {
        for (let c = 0; c < colCount; c++) {
          const cell = worksheet.getCell(bounds.startRow + r, bounds.startCol + c);
          let styleToApply: Record<string, unknown> | null = flatStyle;

          if (args.styles) {
            const gridStyle = args.styles[r]?.[c];
            if (gridStyle === null) continue;
            if (gridStyle) styleToApply = gridStyle;
          }

          if (styleToApply && Object.keys(styleToApply).length > 0) {
            Cell.applyStyle(cell, styleToApply);
          }
        }
      }

      await excel.save();
      return {
        success: true as const,
        message: `Formatted range ${bounds.startRow},${bounds.startCol}:${bounds.endRow},${bounds.endCol} on sheet '${worksheet.name}'`,
      };
    } catch (error) {
      const msg = error instanceof Error ? error.message : String(error);
      throw new ToolError("Failed to format range: " + msg);
    }
  },
});
