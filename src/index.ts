#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { registerAuthTools } from "./tools/auth-tools.js";
import { registerMailTools } from "./tools/mail-tools.js";
import { registerCalendarTools } from "./tools/calendar-tools.js";
import { registerOneDriveTools } from "./tools/onedrive-tools.js";
import { registerSharePointTools } from "./tools/sharepoint-tools.js";
import { registerTeamsTools } from "./tools/teams-tools.js";

const server = new McpServer({
  name: "office365-mcp-server",
  version: "1.0.0",
});

registerAuthTools(server);
registerMailTools(server);
registerCalendarTools(server);
registerOneDriveTools(server);
registerSharePointTools(server);
registerTeamsTools(server);

const transport = new StdioServerTransport();
await server.connect(transport);
console.error("Office 365 MCP server running on stdio");
