import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import type { TeamsMessageLocation, TeamsService } from "../services/teams.js";
import { safeTool } from "../utils/safe-tool.js";
import {
  formatTeamList,
  formatChannelList,
  formatChannelMessages,
  formatChannelReplies,
  formatChatList,
  formatChatMessages,
} from "../formatters/teams.js";
import { formatStoredDownloads } from "../formatters/download.js";
import { formatSize } from "../utils/format.js";
import { DOWNLOAD_MAX_BYTES } from "../constants.js";

interface MessageLocationParams {
  messageId: string;
  chatId?: string;
  teamId?: string;
  channelId?: string;
  parentMessageId?: string;
}

function resolveChannelLocation(
  teamId: string,
  channelId: string,
  params: MessageLocationParams
): TeamsMessageLocation {
  const { messageId, parentMessageId } = params;
  if (parentMessageId) return { kind: "channelReply", teamId, channelId, parentMessageId, messageId };
  return { kind: "channel", teamId, channelId, messageId };
}

function resolveMessageLocation(params: MessageLocationParams): TeamsMessageLocation {
  if (params.chatId) return { kind: "chat", chatId: params.chatId, messageId: params.messageId };
  if (params.teamId && params.channelId) return resolveChannelLocation(params.teamId, params.channelId, params);
  throw new Error("Informe chatId, ou teamId junto com channelId, para localizar a mensagem.");
}

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
    "Lista mensagens recentes de um canal do Teams (sem as respostas; use list-channel-message-replies para a thread).",
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
    "list-channel-message-replies",
    "Lista as respostas da thread de uma mensagem de canal do Teams.",
    {
      teamId: z.string().describe("ID do time"),
      channelId: z.string().describe("ID do canal"),
      messageId: z.string().describe("ID da mensagem principal da thread"),
      top: z.number().optional().describe("Número máximo de respostas (padrão: 20)"),
    },
    safeTool(async (params) => {
      const replies = await teams.listChannelMessageReplies(params);
      return {
        content: [{ type: "text" as const, text: formatChannelReplies(replies) }],
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
    "list-chat-messages",
    "Lista mensagens recentes de um chat/DM do Teams.",
    {
      chatId: z.string().describe("ID do chat"),
      top: z.number().optional().describe("Número máximo de mensagens (padrão: 20)"),
    },
    safeTool(async (params) => {
      const messages = await teams.listChatMessages(params);
      return {
        content: [{ type: "text" as const, text: formatChatMessages(messages) }],
      };
    })
  );

  server.tool(
    "download-teams-message-files",
    `Baixa arquivos anexados e imagens embutidas de uma mensagem do Teams (chat ou canal) para o disco local e retorna o caminho de cada arquivo, para anexar na resposta. Informe chatId para mensagem de chat/DM, ou teamId + channelId para mensagem de canal (mais parentMessageId se for resposta em thread). Sem attachmentName, baixa todos os arquivos anexados e as imagens embutidas no corpo. Limite: ${formatSize(DOWNLOAD_MAX_BYTES)} por arquivo.`,
    {
      messageId: z.string().describe("ID da mensagem"),
      chatId: z.string().optional().describe("ID do chat (mensagem de chat/DM)"),
      teamId: z.string().optional().describe("ID do time (mensagem de canal, junto com channelId)"),
      channelId: z.string().optional().describe("ID do canal (mensagem de canal, junto com teamId)"),
      parentMessageId: z.string().optional().describe("ID da mensagem principal quando messageId é uma resposta de thread de canal"),
      attachmentName: z.string().optional().describe("Nome ou parte do nome do arquivo anexado, como aparece em list-chat-messages/list-channel-messages/list-channel-message-replies"),
    },
    safeTool(async (params) => {
      const location = resolveMessageLocation(params);
      const downloads = await teams.downloadMessageFiles(location, params.attachmentName);
      return {
        content: [{ type: "text" as const, text: formatStoredDownloads(downloads) }],
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
