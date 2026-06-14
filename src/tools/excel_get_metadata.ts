import { defineTool, ToolError } from "@achmadya-dev/mcp-core";
import { z } from "zod";
import { Excel } from "../excel/excel.js";
import { Range } from "../excel/range.js";

const filePathField = z.string().describe("Absolute path to the local Excel file (.xlsx)");

const getMetadataInputSchema = z.object({
  filePath: filePathField,
  includeRanges: z.boolean().optional().describe("Include usedRange and mergedCells per sheet"),
});

const getMetadataOutputShape = z.object({
  filePath: z.string(),
  creator: z.string().nullable(),
  lastModifiedBy: z.string().nullable(),
  created: z.string().nullable(),
  modified: z.string().nullable(),
  sheets: z.array(
    z.object({
      name: z.string(),
      rowCount: z.number(),
      columnCount: z.number(),
      state: z.enum(["visible", "hidden", "veryHidden"]).optional(),
      usedRange: z.string().nullable().optional(),
      mergedCells: z.array(z.string()).optional(),
    })
  ),
});

export const excel_get_metadata = defineTool({
  name: "excel_get_metadata",
  description:
    "Workbook metadata: sheet names, dimensions, creator, dates. Use includeRanges for usedRange and mergedCells per sheet (replaces list_sheets).",
  inputSchema: getMetadataInputSchema,
  outputSchema: getMetadataOutputShape,
  handler: async (args) => {
    try {
      const { filePath, includeRanges = false } = args;
      const excel = await Excel.open(filePath);
      const workbook = excel.workbook;

      const sheets = workbook.worksheets.map((ws) => {
        const base = {
          name: ws.name,
          rowCount: ws.rowCount,
          columnCount: ws.actualColumnCount || 0,
          state: ws.state ?? "visible",
        };
        if (!includeRanges) return base;

        let usedRange: string | null = null;
        if (ws.rowCount > 0 && (ws.actualColumnCount || 0) > 0) {
          usedRange = Range.toA1({
            startRow: 1,
            startCol: 1,
            endRow: ws.rowCount,
            endCol: ws.actualColumnCount || 1,
          });
        }

        const mergedCells = (ws.model.merges ?? []).map((m) => String(m));
        return { ...base, usedRange, mergedCells };
      });

      return {
        filePath,
        creator: workbook.creator || null,
        lastModifiedBy: workbook.lastModifiedBy || null,
        created: workbook.created ? workbook.created.toISOString() : null,
        modified: workbook.modified ? workbook.modified.toISOString() : null,
        sheets,
      };
    } catch (error) {
      const msg = error instanceof Error ? error.message : String(error);
      throw new ToolError("Failed to get metadata: " + msg);
    }
  },
});
