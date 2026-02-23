import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import * as teams from "../services/teams.js";
import { safeTool } from "../utils/errors.js";

export function registerTeamsTools(server: McpServer) {
  server.tool(
    "list-teams",
    "Lista os times do Microsoft Teams que o usuário participa.",
    {},
    safeTool(async () => {
      const teamsList = await teams.listTeams();
      const formatted = teamsList.map(
        (t: any) => `- **${t.displayName}**\n  ${t.description ?? ""}\n  ID: ${t.id}`
      );

      return {
        content: [
          {
            type: "text" as const,
            text: formatted.length > 0
              ? `## Times (${formatted.length})\n\n${formatted.join("\n\n")}`
              : "Nenhum time encontrado.",
          },
        ],
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
      const formatted = channels.map(
        (c: any) =>
          `- **${c.displayName}** (${c.membershipType})\n  ${c.description ?? ""}\n  ID: ${c.id}`
      );

      return {
        content: [
          {
            type: "text" as const,
            text: formatted.length > 0
              ? `## Canais (${formatted.length})\n\n${formatted.join("\n\n")}`
              : "Nenhum canal encontrado.",
          },
        ],
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
      const formatted = messages.map((m: any) => {
        const from = m.from?.user?.displayName ?? "Desconhecido";
        const date = new Date(m.createdDateTime).toLocaleString("pt-BR");
        const content = m.body?.content?.substring(0, 200) ?? "";
        return `- **${from}** (${date})\n  ${content}`;
      });

      return {
        content: [
          {
            type: "text" as const,
            text: formatted.length > 0
              ? `## Mensagens do Canal (${formatted.length})\n\n${formatted.join("\n\n")}`
              : "Nenhuma mensagem encontrada.",
          },
        ],
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
      const formatted = chats.map((c: any) => {
        const topic = c.topic ?? "Chat sem título";
        const type = c.chatType ?? "unknown";
        const updated = c.lastUpdatedDateTime
          ? new Date(c.lastUpdatedDateTime).toLocaleString("pt-BR")
          : "N/A";
        const members =
          c.members?.map((m: any) => m.displayName).join(", ") ?? "";
        return `- **${topic}** (${type})\n  Membros: ${members}\n  Última atualização: ${updated}\n  ID: ${c.id}`;
      });

      return {
        content: [
          {
            type: "text" as const,
            text: formatted.length > 0
              ? `## Chats (${formatted.length})\n\n${formatted.join("\n\n")}`
              : "Nenhum chat encontrado.",
          },
        ],
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
