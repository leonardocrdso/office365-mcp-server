import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import type { TeamsService } from "../services/teams.js";
import { safeTool } from "../utils/errors.js";
import {
  formatTeamList,
  formatChannelList,
  formatChannelMessages,
  formatChatList,
} from "../formatters/teams.js";

export function registerTeamsTools(server: McpServer, teams: TeamsService) {
  server.tool(
    "list-teams",
    "Lista os times do Microsoft Teams que o usuário participa.",
    {},
    safeTool(async () => {
      const teamsList = await teams.listTeams();
      return {
        content: [{ type: "text" as const, text: formatTeamList(teamsList) }],
      };
    })
  );

  server.tool(
    "list-channels",
    "Lista canais de um time do Teams.",
    {
      teamId: z.string().describe("ID do time"),
    },
    safeTool(async (params) => {
      const channels = await teams.listChannels(params.teamId);
      return {
        content: [{ type: "text" as const, text: formatChannelList(channels) }],
      };
    })
  );

  server.tool(
    "list-channel-messages",
    "Lista mensagens recentes de um canal do Teams.",
    {
      teamId: z.string().describe("ID do time"),
      channelId: z.string().describe("ID do canal"),
      top: z.number().optional().describe("Número máximo de mensagens (padrão: 20)"),
    },
    safeTool(async (params) => {
      const messages = await teams.listChannelMessages(params);
      return {
        content: [{ type: "text" as const, text: formatChannelMessages(messages) }],
      };
    })
  );

  server.tool(
    "send-channel-message",
    "Envia uma mensagem em um canal do Teams.",
    {
      teamId: z.string().describe("ID do time"),
      channelId: z.string().describe("ID do canal"),
      content: z.string().describe("Conteúdo da mensagem"),
      contentType: z.enum(["text", "html"]).optional().describe("Tipo de conteúdo (padrão: text)"),
    },
    safeTool(async (params) => {
      await teams.sendChannelMessage(params);
      return {
        content: [
          { type: "text" as const, text: "Mensagem enviada no canal com sucesso." },
        ],
      };
    })
  );

  server.tool(
    "list-chats",
    "Lista conversas/chats diretos do Teams.",
    {
      top: z.number().optional().describe("Número máximo de chats (padrão: 20)"),
    },
    safeTool(async (params) => {
      const chats = await teams.listChats(params.top);
      return {
        content: [{ type: "text" as const, text: formatChatList(chats) }],
      };
    })
  );

  server.tool(
    "send-chat-message",
    "Envia uma mensagem em um chat direto do Teams.",
    {
      chatId: z.string().describe("ID do chat"),
      content: z.string().describe("Conteúdo da mensagem"),
      contentType: z.enum(["text", "html"]).optional().describe("Tipo de conteúdo (padrão: text)"),
    },
    safeTool(async (params) => {
      await teams.sendChatMessage(params);
      return {
        content: [
          { type: "text" as const, text: "Mensagem enviada no chat com sucesso." },
        ],
      };
    })
  );
}
