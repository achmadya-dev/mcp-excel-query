#!/usr/bin/env node
import { Server } from "./mcp/server.js";
import packageJson from "../package.json" with { type: "json" };
import { allExcelTools } from "./mcp/registry.js";

async function main(): Promise<void> {
  const server = new Server({
    name: "Excel Local Manager",
    version: packageJson.version,
  });

  for (const tool of allExcelTools) {
    server.registerTool(tool as unknown as Parameters<Server["registerTool"]>[0]);
  }

  await server.start();
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
