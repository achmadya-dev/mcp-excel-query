# @achmadya-dev/mcp-excel-query

MCP server for local `.xlsx` files over **stdio**. TypeScript + [ExcelJS](https://github.com/exceljs/exceljs). Read, write, format, and manage workbooks without Microsoft Excel installed.

## Requirements

- Node.js **≥ 20**
- `.xlsx` files on the local filesystem (paths passed in tool arguments)

## Install from npm

```json
{
  "mcpServers": {
    "excel": {
      "command": "npx",
      "args": ["-y", "@achmadya-dev/mcp-excel-query"],
      "env": {
        "EXCEL_PAGING_CELLS_LIMIT": "4000"
      }
    }
  }
}
```

Or use `envFile` instead of inline `env`.

## Develop from source

```bash
cp .env.example .env
pnpm install
pnpm --filter @achmadya-dev/mcp-excel-query run build
```

`.cursor/mcp.json`:

```json
{
  "mcpServers": {
    "excel": {
      "command": "node",
      "args": ["${workspaceFolder}/packages/mcp-excel-query/dist/index.js"],
      "envFile": "${workspaceFolder}/.env"
    }
  }
}
```

Relevant `.env` key:

```env
EXCEL_PAGING_CELLS_LIMIT=4000
```

## Environment variables

| Variable                   | Default              | Description                                            |
| -------------------------- | -------------------- | ------------------------------------------------------ |
| `EXCEL_PAGING_CELLS_LIMIT` | _(unset = no limit)_ | Max cells returned per read; prevents huge sheet dumps |

## Tools (17)

| Tool                         | Use for                                                                      |
| ---------------------------- | ---------------------------------------------------------------------------- |
| `excel_read_sheet`           | Read rows, formulas (`showFormula`), styles, dates, merged cells, validation |
| `excel_get_metadata`         | Sheet names, dimensions, date info                                           |
| `excel_create_file`          | New workbook + optional headers                                              |
| `excel_write_range`          | Bulk write, append, new sheet, single cell/formula, optional style           |
| `excel_format_range`         | Font, fill, border, alignment, numFmt, merge                                 |
| `excel_copy_sheet`           | Duplicate worksheet                                                          |
| `excel_rename_sheet`         | Rename worksheet                                                             |
| `excel_delete_sheet`         | Delete worksheet                                                             |
| `excel_copy_range`           | Copy cells to another location                                               |
| `excel_delete_range`         | Delete range; `shiftDirection: up` or `left`                                 |
| `excel_unmerge_cells`        | Unmerge range                                                                |
| `excel_create_table`         | Native Excel table                                                           |
| `excel_insert_rows`          | Insert empty rows                                                            |
| `excel_insert_columns`       | Insert empty columns                                                         |
| `excel_set_sheet_visibility` | Show, hide, or very-hide a sheet                                             |
| `excel_set_data_validation`  | Dropdown / validation rules on a range                                       |
| `excel_set_dimensions`       | Row height, column width, sheet defaults                                     |

### Common patterns

```json
// Append rows
{ "append": true, "values": [[1, "Alice", 25]] }

// Write formula to one cell
{ "startCell": "D2", "values": [["=SUM(A2:C2)"]] }

// New blank sheet
{ "newSheet": true, "sheetName": "Report" }

// Style one cell
{ "startCell": "A1", "values": [["Title"]], "style": { "font": { "bold": true } } }

// Merge + format
{ "startCell": "A1", "endCell": "C1", "mergeCells": true, "bold": true }

// Read dates using each cell's Excel numFmt
{ "range": "H3:H10", "headerRow": 0, "dateFormat": "cell" }

// Hide a sheet
{ "sheetName": "Archive", "state": "hidden" }

// Dropdown validation
{ "startCell": "B2", "endCell": "B100", "type": "list", "formulae": ["\"Yes,No,Maybe\""] }
```

## Notes

- Formulas are stored, not evaluated server-side.
- `.xlsx` only (not `.xls` or `.csv`).

## Package scripts

```bash
pnpm run build
pnpm test
pnpm start
```
