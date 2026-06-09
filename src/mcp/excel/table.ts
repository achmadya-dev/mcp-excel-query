import {
  loadWorkbook,
  operationSuccess,
  parseA1Range,
  resolveWorksheet,
  saveWorkbook,
} from "./utils.js";

export async function createTable(params: {
  filePath: string;
  sheetName?: string;
  dataRange: string;
  tableName?: string;
  tableStyle?: string;
}) {
  const workbook = await loadWorkbook(params.filePath);
  const worksheet = resolveWorksheet(workbook, params.sheetName);
  const bounds = parseA1Range(params.dataRange);

  const headerRow = worksheet.getRow(bounds.startRow);
  const columns: { name: string; filterButton?: boolean }[] = [];
  for (let c = bounds.startCol; c <= bounds.endCol; c++) {
    const val = headerRow.getCell(c).value;
    const name =
      val !== null && val !== undefined && val !== ""
        ? String(val)
        : `Column${c - bounds.startCol + 1}`;
    columns.push({ name, filterButton: true });
  }

  const rows: unknown[][] = [];
  for (let r = bounds.startRow + 1; r <= bounds.endRow; r++) {
    const row = worksheet.getRow(r);
    const rowData: unknown[] = [];
    for (let c = bounds.startCol; c <= bounds.endCol; c++) {
      rowData.push(row.getCell(c).value ?? null);
    }
    rows.push(rowData);
  }

  worksheet.addTable({
    name: params.tableName ?? `Table_${Date.now()}`,
    ref: params.dataRange,
    headerRow: true,
    style: {
      theme: (params.tableStyle ?? "TableStyleMedium9") as "TableStyleMedium9",
      showRowStripes: true,
    },
    columns,
    rows,
  });

  await saveWorkbook(workbook, params.filePath);
  return operationSuccess(
    `Table '${params.tableName ?? "created"}' added on sheet '${worksheet.name}' at ${params.dataRange}`
  );
}
