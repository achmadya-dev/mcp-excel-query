import { defineTool, ToolError } from "@achmadya-dev/mcp-core";
import { z } from "zod";
import type ExcelJS from "exceljs";
import type { DataValidation } from "exceljs";
import { Excel } from "../excel/excel.js";
import { Range } from "../excel/range.js";

const filePathField = z.string().describe("Absolute path to the local Excel file (.xlsx)");

const sheetNameOptional = z
  .string()
  .optional()
  .describe("Name of the worksheet (defaults to the first sheet)");

const validationOperatorSchema = z.enum([
  "between",
  "notBetween",
  "equal",
  "notEqual",
  "greaterThan",
  "lessThan",
  "greaterThanOrEqual",
  "lessThanOrEqual",
]);

const setDataValidationInputSchema = z.object({
  filePath: filePathField,
  sheetName: sheetNameOptional,
  startCell: z.string(),
  endCell: z.string().optional(),
  range: z.string().optional(),
  clear: z.boolean().optional().describe("Remove data validation from the range"),
  type: z
    .enum(["list", "whole", "decimal", "date", "textLength", "custom"])
    .optional()
    .describe("Validation type (required unless clear:true)"),
  operator: validationOperatorSchema.optional(),
  formulae: z
    .array(z.string())
    .optional()
    .describe('Formula strings, e.g. ["\\"Yes,No\\""] for list or ["=A1>0"] for custom'),
  allowBlank: z.boolean().optional(),
  showInputMessage: z.boolean().optional(),
  prompt: z.string().optional(),
  promptTitle: z.string().optional(),
  showErrorMessage: z.boolean().optional(),
  error: z.string().optional(),
  errorTitle: z.string().optional(),
});

const operationResultOutputShape = z.object({
  success: z.boolean(),
  message: z.string(),
});

type ValidationOperator =
  | "between"
  | "notBetween"
  | "equal"
  | "notEqual"
  | "greaterThan"
  | "lessThan"
  | "greaterThanOrEqual"
  | "lessThanOrEqual";

type ValidationType = "list" | "whole" | "decimal" | "date" | "textLength" | "custom";

type WorksheetWithValidations = ExcelJS.Worksheet & {
  dataValidations: {
    remove: (address: string) => void;
    model: Record<string, unknown>;
  };
};

function boundsFromParams(params: { range?: string; startCell: string; endCell?: string }) {
  return params.range ? Range.parse(params.range) : Range.between(params.startCell, params.endCell);
}

function applyValidationToCell(
  cell: ExcelJS.Cell,
  params: {
    type: ValidationType;
    operator?: ValidationOperator;
    formulae?: string[];
    allowBlank?: boolean;
    showInputMessage?: boolean;
    prompt?: string;
    promptTitle?: string;
    showErrorMessage?: boolean;
    error?: string;
    errorTitle?: string;
  }
): void {
  const validation: DataValidation = {
    type: params.type,
    formulae: params.formulae ?? [],
    ...(params.operator ? { operator: params.operator } : {}),
    ...(params.allowBlank !== undefined ? { allowBlank: params.allowBlank } : {}),
    ...(params.showInputMessage !== undefined ? { showInputMessage: params.showInputMessage } : {}),
    ...(params.prompt ? { prompt: params.prompt } : {}),
    ...(params.promptTitle ? { promptTitle: params.promptTitle } : {}),
    ...(params.showErrorMessage !== undefined ? { showErrorMessage: params.showErrorMessage } : {}),
    ...(params.error ? { error: params.error } : {}),
    ...(params.errorTitle ? { errorTitle: params.errorTitle } : {}),
  };
  cell.dataValidation = validation;
}

export const excel_set_data_validation = defineTool({
  name: "excel_set_data_validation",
  description:
    "Set or clear data validation on a range. Supports list, whole, decimal, date, textLength, and custom rules with formulae.",
  inputSchema: setDataValidationInputSchema,
  outputSchema: operationResultOutputShape,
  handler: async (args) => {
    try {
      const excel = await Excel.open(args.filePath);
      const sheet = excel.sheet(args.sheetName);
      const worksheet = sheet.raw;
      const bounds = boundsFromParams(args);

      for (let r = bounds.startRow; r <= bounds.endRow; r++) {
        for (let c = bounds.startCol; c <= bounds.endCol; c++) {
          const cell = worksheet.getCell(r, c);
          if (args.clear) {
            const address = cell.address;
            const validations = (worksheet as WorksheetWithValidations).dataValidations;
            validations.remove(address);
            delete validations.model[address];
            continue;
          }

          if (!args.type) {
            throw new Error("type is required when setting data validation");
          }
          if (!args.formulae || args.formulae.length === 0) {
            throw new Error("formulae is required when setting data validation");
          }

          applyValidationToCell(cell, {
            type: args.type,
            operator: args.operator,
            formulae: args.formulae,
            allowBlank: args.allowBlank,
            showInputMessage: args.showInputMessage,
            prompt: args.prompt,
            promptTitle: args.promptTitle,
            showErrorMessage: args.showErrorMessage,
            error: args.error,
            errorTitle: args.errorTitle,
          });
        }
      }

      await excel.save();

      const rangeLabel = args.range ?? `${args.startCell}${args.endCell ? `:${args.endCell}` : ""}`;
      if (args.clear) {
        return {
          success: true as const,
          message: `Cleared data validation on sheet '${worksheet.name}' at ${rangeLabel}`,
        };
      }
      return {
        success: true as const,
        message: `Set ${args.type} data validation on sheet '${worksheet.name}' at ${rangeLabel}`,
      };
    } catch (error) {
      const msg = error instanceof Error ? error.message : String(error);
      throw new ToolError("Failed to set data validation: " + msg);
    }
  },
});
