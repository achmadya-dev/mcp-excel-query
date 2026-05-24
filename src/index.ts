#!/usr/bin/env node
import { Server } from "./mcp/server.js";
import packageJson from "../package.json" with { type: "json" };
import {
  excel_list_sheets,
  excel_read_sheet,
  excel_get_metadata,
  excel_create_file,
  excel_create_sheet,
  excel_write_cell,
  excel_append_rows,
} from "./mcp/registry.js";

async function main(): Promise<void> {
  const server = new Server({
    name: "Excel Local Manager",
    version: packageJson.version,
  });

  server.registerTool(excel_list_sheets);
  server.registerTool(excel_read_sheet);
  server.registerTool(excel_get_metadata);
  server.registerTool(excel_create_file);
  server.registerTool(excel_create_sheet);
  server.registerTool(excel_write_cell);
  server.registerTool(excel_append_rows);

  await server.start();
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
