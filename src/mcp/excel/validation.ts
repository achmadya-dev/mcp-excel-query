import ExcelJS from "exceljs";
import type { DataValidation } from "exceljs";
import {
  loadWorkbook,
  operationSuccess,
  parseA1Range,
  resolveRangeBounds,
  resolveWorksheet,
  saveWorkbook,
} from "./utils.js";

type ValidationOperator =
  | "between"
  | "notBetween"
  | "equal"
  | "notEqual"
  | "greaterThan"
  | "lessThan"
  | "greaterThanOrEqual"
  | "lessThanOrEqual";

type ValidationType =
  | "list"
  | "whole"
  | "decimal"
  | "date"
  | "textLength"
  | "custom";

function resolveBounds(params: {
  range?: string;
  startCell: string;
  endCell?: string;
}) {
  return params.range
    ? parseA1Range(params.range)
    : resolveRangeBounds(params.startCell, params.endCell);
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
    ...(params.showInputMessage !== undefined
      ? { showInputMessage: params.showInputMessage }
      : {}),
    ...(params.prompt ? { prompt: params.prompt } : {}),
    ...(params.promptTitle ? { promptTitle: params.promptTitle } : {}),
    ...(params.showErrorMessage !== undefined
      ? { showErrorMessage: params.showErrorMessage }
      : {}),
    ...(params.error ? { error: params.error } : {}),
    ...(params.errorTitle ? { errorTitle: params.errorTitle } : {}),
  };
  cell.dataValidation = validation;
}

type WorksheetWithValidations = ExcelJS.Worksheet & {
  dataValidations: {
    remove: (address: string) => void;
    model: Record<string, unknown>;
  };
};

export async function setDataValidation(params: {
  filePath: string;
  sheetName?: string;
  startCell: string;
  endCell?: string;
  range?: string;
  clear?: boolean;
  type?: ValidationType;
  operator?: ValidationOperator;
  formulae?: string[];
  allowBlank?: boolean;
  showInputMessage?: boolean;
  prompt?: string;
  promptTitle?: string;
  showErrorMessage?: boolean;
  error?: string;
  errorTitle?: string;
}) {
  const workbook = await loadWorkbook(params.filePath);
  const worksheet = resolveWorksheet(workbook, params.sheetName);
  const bounds = resolveBounds(params);

  for (let r = bounds.startRow; r <= bounds.endRow; r++) {
    for (let c = bounds.startCol; c <= bounds.endCol; c++) {
      const cell = worksheet.getCell(r, c);
      if (params.clear) {
        const address = cell.address;
        const validations = (worksheet as WorksheetWithValidations).dataValidations;
        validations.remove(address);
        delete validations.model[address];
        continue;
      }

      if (!params.type) {
        throw new Error("type is required when setting data validation");
      }
      if (!params.formulae || params.formulae.length === 0) {
        throw new Error("formulae is required when setting data validation");
      }

      applyValidationToCell(cell, {
        type: params.type,
        operator: params.operator,
        formulae: params.formulae,
        allowBlank: params.allowBlank,
        showInputMessage: params.showInputMessage,
        prompt: params.prompt,
        promptTitle: params.promptTitle,
        showErrorMessage: params.showErrorMessage,
        error: params.error,
        errorTitle: params.errorTitle,
      });
    }
  }

  await saveWorkbook(workbook, params.filePath);

  const rangeLabel = params.range ?? `${params.startCell}${params.endCell ? `:${params.endCell}` : ""}`;
  if (params.clear) {
    return operationSuccess(
      `Cleared data validation on sheet '${worksheet.name}' at ${rangeLabel}`
    );
  }
  return operationSuccess(
    `Set ${params.type} data validation on sheet '${worksheet.name}' at ${rangeLabel}`
  );
}
