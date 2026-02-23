import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { registerAuthTools } from "./tools/auth-tools.ts";
import { registerMailTools } from "./tools/mail-tools.ts";
import { registerCalendarTools } from "./tools/calendar-tools.ts";
import { registerOneDriveTools } from "./tools/onedrive-tools.ts";
import { registerSharePointTools } from "./tools/sharepoint-tools.ts";
import { registerTeamsTools } from "./tools/teams-tools.ts";

const server = new McpServer({
  name: "office365",
  version: "1.0.0",
});

// Registrar todas as tools
registerAuthTools(server);
registerMailTools(server);
registerCalendarTools(server);
registerOneDriveTools(server);
registerSharePointTools(server);
registerTeamsTools(server);

// Conectar via stdio (para integração com Claude Code)
const transport = new StdioServerTransport();
await server.connect(transport);
