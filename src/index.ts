#!/usr/bin/env node
import { startMcpServer } from "@achmadya-dev/mcp-core";
import packageJson from "../package.json" with { type: "json" };
import { excel_copy_range } from "./tools/excel_copy_range.js";
import { excel_copy_sheet } from "./tools/excel_copy_sheet.js";
import { excel_create_file } from "./tools/excel_create_file.js";
import { excel_create_table } from "./tools/excel_create_table.js";
import { excel_delete_range } from "./tools/excel_delete_range.js";
import { excel_delete_sheet } from "./tools/excel_delete_sheet.js";
import { excel_format_range } from "./tools/excel_format_range.js";
import { excel_get_metadata } from "./tools/excel_get_metadata.js";
import { excel_insert_columns } from "./tools/excel_insert_columns.js";
import { excel_insert_rows } from "./tools/excel_insert_rows.js";
import { excel_read_sheet } from "./tools/excel_read_sheet.js";
import { excel_rename_sheet } from "./tools/excel_rename_sheet.js";
import { excel_set_data_validation } from "./tools/excel_set_data_validation.js";
import { excel_set_dimensions } from "./tools/excel_set_dimensions.js";
import { excel_set_sheet_visibility } from "./tools/excel_set_sheet_visibility.js";
import { excel_unmerge_cells } from "./tools/excel_unmerge_cells.js";
import { excel_write_range } from "./tools/excel_write_range.js";

await startMcpServer({
  name: "Excel Local Manager",
  version: packageJson.version,
  tools: [
    excel_read_sheet,
    excel_get_metadata,
    excel_create_file,
    excel_write_range,
    excel_format_range,
    excel_copy_sheet,
    excel_rename_sheet,
    excel_delete_sheet,
    excel_copy_range,
    excel_delete_range,
    excel_unmerge_cells,
    excel_create_table,
    excel_insert_rows,
    excel_insert_columns,
    excel_set_sheet_visibility,
    excel_set_data_validation,
    excel_set_dimensions,
  ],
});
