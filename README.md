# mcp-excel-query

MCP server for local `.xlsx` files via **stdio**. TypeScript + [ExcelJS](https://github.com/exceljs/exceljs).

## Install

```json
{
  "mcpServers": {
    "excel": {
      "command": "npx",
      "args": ["-y", "@achmadya-dev/mcp-excel-query"],
      "env": { "EXCEL_PAGING_CELLS_LIMIT": "4000" }
    }
  }
}
```

## Tools (17)

| Tool | Use for |
|------|---------|
| `excel_read_sheet` | Read JSON rows, formulas (`showFormula`), styles, `dateFormat` / `dateFormat: "cell"`, merged cells, validation metadata |
| `excel_get_metadata` | Sheet names, dimensions, dates — use instead of a separate list-sheets tool |
| `excel_create_file` | New workbook + optional headers |
| `excel_write_range` | **Main write tool**: bulk data, append (`append:true`), new blank sheet (`newSheet:true`), single cell/formula (`startCell` + `values`), optional `style` |
| `excel_format_range` | Font, fill, border, alignment, numFmt, merge (`mergeCells:true`) |
| `excel_copy_sheet` | Duplicate worksheet |
| `excel_rename_sheet` | Rename worksheet |
| `excel_delete_sheet` | Delete worksheet |
| `excel_copy_range` | Copy cells to another location |
| `excel_delete_range` | Delete range; `shiftDirection: up` = rows, `left` = columns |
| `excel_unmerge_cells` | Unmerge range |
| `excel_create_table` | Native Excel table |
| `excel_insert_rows` | Insert empty rows |
| `excel_insert_columns` | Insert empty columns |
| `excel_set_sheet_visibility` | Show, hide, or very-hide a sheet |
| `excel_set_data_validation` | Set or clear dropdown/rules validation on a range |
| `excel_set_dimensions` | Row height, column width, sheet defaults |

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

// Read dates in human-readable form (uses each cell's Excel numFmt)
{ "range": "H3:H10", "headerRow": 0, "dateFormat": "cell" }

// Write dates as text — auto-parsed to Excel date values
{ "startCell": "H3", "values": [["09 Jun 2026"]] }

// Hide a sheet
{ "sheetName": "Archive", "state": "hidden" }

// Dropdown list validation
{ "startCell": "B2", "endCell": "B100", "type": "list", "formulae": ["\"Yes,No,Maybe\""] }

// Column width and row height
{ "range": "A1:D20", "columnWidth": 18, "rowHeight": 22 }
```

## Notes

- Formulas are stored, not evaluated server-side.
- `.xlsx` only.

## Development

```bash
pnpm install && pnpm run build && pnpm test
```
