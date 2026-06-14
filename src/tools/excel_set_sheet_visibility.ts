import { defineTool, ToolError } from "@achmadya-dev/mcp-core";
import { z } from "zod";
import { Excel } from "../excel/excel.js";

const filePathField = z.string().describe("Absolute path to the local Excel file (.xlsx)");

const setSheetVisibilityInputSchema = z.object({
  filePath: filePathField,
  sheetName: z.string().describe("Worksheet to show or hide"),
  state: z
    .enum(["visible", "hidden", "veryHidden"])
    .describe(
      "visible = shown; hidden = hidden (user can unhide); veryHidden = hidden from Excel UI unhide menu"
    ),
});

const operationResultOutputShape = z.object({
  success: z.boolean(),
  message: z.string(),
});

export const excel_set_sheet_visibility = defineTool({
  name: "excel_set_sheet_visibility",
  description: "Show, hide, or very-hide a worksheet (state: visible, hidden, veryHidden).",
  inputSchema: setSheetVisibilityInputSchema,
  outputSchema: operationResultOutputShape,
  handler: async (args) => {
    try {
      const { filePath, sheetName, state } = args;
      const excel = await Excel.open(filePath);
      const sheet = excel.sheet(sheetName);
      sheet.raw.state = state;
      await excel.save();
      return {
        success: true as const,
        message: `Sheet '${sheetName}' visibility set to '${state}'`,
      };
    } catch (error) {
      const msg = error instanceof Error ? error.message : String(error);
      throw new ToolError("Failed to set sheet visibility: " + msg);
    }
  },
});
