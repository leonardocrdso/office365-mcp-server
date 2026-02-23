#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { msalClient } from "./auth/msal-client.js";
import { createMailService } from "./services/mail.js";
import { createCalendarService } from "./services/calendar.js";
import { createOneDriveService } from "./services/onedrive.js";
import { createSharePointService } from "./services/sharepoint.js";
import { createTeamsService } from "./services/teams.js";
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

const mailService = createMailService(msalClient);
const calendarService = createCalendarService(msalClient);
const oneDriveService = createOneDriveService(msalClient);
const sharePointService = createSharePointService(msalClient);
const teamsService = createTeamsService(msalClient);

registerAuthTools(server);
registerMailTools(server, mailService);
registerCalendarTools(server, calendarService);
registerOneDriveTools(server, oneDriveService);
registerSharePointTools(server, sharePointService);
registerTeamsTools(server, teamsService);

const transport = new StdioServerTransport();
await server.connect(transport);
console.error("Office 365 MCP server running on stdio");
