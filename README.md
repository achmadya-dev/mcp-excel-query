# mcp-excel-query

Model Context Protocol (MCP) server for local Excel file (.xlsx) reading and manipulation. It lets MCP clients safely query, edit, and style local spreadsheets via stdio.

## Requirements

- Node.js **≥ 20**

Communication uses **stdio** (not HTTP).

## Install in MCP Clients (e.g. Cursor)

1. Open **Settings → MCP**, or edit the `mcp.json` file for your Cursor account.
2. Add a server entry like the example below. The `npx -y` command fetches the package from the npm registry and runs it (no global install required).

```json
{
  "mcpServers": {
    "excel": {
      "command": "npx",
      "args": ["-y", "@achmadya-dev/mcp-excel-query"]
    }
  }
}
```

## Manual setup from a cloned repository

Clone the repository, install dependencies, then build:

```bash
git clone <repo-url> mcp-excel-query
cd mcp-excel-query
pnpm install && pnpm run build
```

Then register the MCP server with **`node`** and the **absolute path** to `dist/index.js` in your project folder:

```json
{
  "mcpServers": {
    "excel": {
      "command": "node",
      "args": ["/home/madya/Repository/RND/mcp/mcp-excel-query/dist/index.js"]
    }
  }
}
```

Replace the path in `args` with your clone location. After changing TypeScript sources, run `pnpm run build` again.

## Tools

The server registers 7 distinct tools:

### Reading & Navigation

1. **`excel_list_sheets`**
   - List the names of all worksheets inside an Excel file (.xlsx).
   - Parameters:
     - `filePath` (string, required): Absolute path to the local Excel file.
2. **`excel_read_sheet`**
   - Read data rows from a specific worksheet. Supports range (A1 notation), headerRow index, and paginated reading (limit/offset).
   - Parameters:
     - `filePath` (string, required): Absolute path to the local Excel file.
     - `sheetName` (string, optional): Name of the worksheet (defaults to first sheet).
     - `range` (string, optional): Cell range in A1 notation (e.g., `"A1:C10"`).
     - `headerRow` (number, optional): Row number containing headers, 1-indexed (defaults to 1). Set to 0 if there are no headers.
     - `limit` (number, optional): Max data rows to return (for pagination).
     - `offset` (number, optional): Number of data rows to skip (for pagination).
3. **`excel_get_metadata`**
   - Retrieve metadata from an Excel file, including creator, modification dates, and lists of sheets with their row/column dimensions.
   - Parameters:
     - `filePath` (string, required): Absolute path to the local Excel file.

### Writing & Manipulation

4. **`excel_create_file`**
   - Create a new Excel file (.xlsx) with an optional first worksheet and optional headers.
   - Parameters:
     - `filePath` (string, required): Absolute path where the file will be created.
     - `sheetName` (string, optional): Name of the first sheet (defaults to `"Sheet1"`).
     - `headers` (array of strings, optional): Header column names to write in the first row.
5. **`excel_create_sheet`**
   - Add a new blank worksheet to an existing Excel file (.xlsx).
   - Parameters:
     - `filePath` (string, required): Absolute path to the local Excel file.
     - `sheetName` (string, required): Name of the new worksheet to create.
6. **`excel_write_cell`**
   - Write or update a value in a specific cell (e.g., `'B5'`) of a worksheet and optionally apply cell styles (font, background fill, borders, alignment).
   - Parameters:
     - `filePath` (string, required): Absolute path to the local Excel file.
     - `sheetName` (string, optional): Name of the worksheet (defaults to first sheet).
     - `cell` (string, required): Cell address in A1 notation (e.g., `"B5"`).
     - `value` (string | number | boolean | null, optional): Value to write.
     - `style` (object, optional): Formatting styles (e.g., `font`, `fill`, `alignment`, `border`).
7. **`excel_append_rows`**
   - Append one or more rows of data to the bottom of a worksheet. Each row is represented as an array of cell values.
   - Parameters:
     - `filePath` (string, required): Absolute path to the local Excel file.
     - `sheetName` (string, optional): Name of the worksheet (defaults to first sheet).
     - `rows` (array of arrays, required): Multi-row array of cell values.

## Other behavior

- Values extracted are automatically converted to clean JSON primitives. Formulas return their calculated result value where possible, and rich-text is combined into a single flat string.
- Cell styles under `excel_write_cell` support standard HEX color definitions (automatically converted to ARGB format) and basic styling properties provided by `exceljs`.
