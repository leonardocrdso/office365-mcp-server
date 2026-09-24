import type {
  GraphTeam,
  GraphChannel,
  GraphChatMessage,
  GraphChat,
} from "../types/graph.js";
import { messageSenderName } from "../types/graph.js";
import { extractHostedContentIds } from "../services/teams.js";
import { MESSAGE_CONTENT_MAX_LENGTH } from "../constants.js";
import { formatDateBR } from "../utils/date.js";

export function formatTeamList(teams: GraphTeam[]): string {
  if (teams.length === 0) return "Nenhum time encontrado.";

  const formatted = teams.map(
    (t) => `- **${t.displayName}**\n  ${t.description ?? ""}\n  ID: ${t.id}`
  );

  return `## Times (${formatted.length})\n\n${formatted.join("\n\n")}`;
}

export function formatChannelList(channels: GraphChannel[]): string {
  if (channels.length === 0) return "Nenhum canal encontrado.";

  const formatted = channels.map(
    (c) =>
      `- **${c.displayName}** (${c.membershipType})\n  ${c.description ?? ""}\n  ID: ${c.id}`
  );

  return `## Canais (${formatted.length})\n\n${formatted.join("\n\n")}`;
}

function formatMessage(message: GraphChatMessage): string {
  const from = messageSenderName(message);
  const date = formatDateBR(message.createdDateTime);
  const content = message.body?.content?.substring(0, MESSAGE_CONTENT_MAX_LENGTH) ?? "";
  const lines = [`- **${from}** (${date})`, `  ${content}`, `  ID: ${message.id}`];

  const fileNames = (message.attachments ?? [])
    .filter((a) => a.contentType === "reference" && a.name)
    .map((a) => `'${a.name}'`);
  if (fileNames.length > 0) lines.push(`  Arquivos: ${fileNames.join(", ")}`);

  const imageCount = extractHostedContentIds(message.body?.content ?? "").length;
  if (imageCount > 0) lines.push(`  Imagens embutidas: ${imageCount}`);

  return lines.join("\n");
}

export function formatChannelMessages(messages: GraphChatMessage[]): string {
  if (messages.length === 0) return "Nenhuma mensagem encontrada.";

  return `## Mensagens do Canal (${messages.length})\n\n${messages.map(formatMessage).join("\n\n")}`;
}

export function formatChannelReplies(replies: GraphChatMessage[]): string {
  if (replies.length === 0) return "Nenhuma resposta encontrada.";

  return `## Respostas da Thread (${replies.length})\n\n${replies.map(formatMessage).join("\n\n")}`;
}

export function formatChatMessages(messages: GraphChatMessage[]): string {
  if (messages.length === 0) return "Nenhuma mensagem encontrada.";

  return `## Mensagens do Chat (${messages.length})\n\n${messages.map(formatMessage).join("\n\n")}`;
}

export function formatChatList(chats: GraphChat[]): string {
  if (chats.length === 0) return "Nenhum chat encontrado.";

  const formatted = chats.map((c) => {
    const topic = c.topic ?? "Chat sem título";
    const type = c.chatType ?? "unknown";
    const updated = c.lastUpdatedDateTime
      ? formatDateBR(c.lastUpdatedDateTime)
      : "N/A";
    const members =
      c.members?.map((m) => m.displayName).join(", ") ?? "";
    return `- **${topic}** (${type})\n  Membros: ${members}\n  Última atualização: ${updated}\n  ID: ${c.id}`;
  });

  return `## Chats (${formatted.length})\n\n${formatted.join("\n\n")}`;
}
