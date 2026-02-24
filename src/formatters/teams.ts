import type {
  GraphTeam,
  GraphChannel,
  GraphChatMessage,
  GraphChat,
} from "../types/graph.js";
import { messageSenderName } from "../types/graph.js";
import { MESSAGE_CONTENT_MAX_LENGTH } from "../constants.js";

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

export function formatChannelMessages(messages: GraphChatMessage[]): string {
  if (messages.length === 0) return "Nenhuma mensagem encontrada.";

  const formatted = messages.map((m) => {
    const from = messageSenderName(m);
    const date = new Date(m.createdDateTime).toLocaleString("pt-BR", { timeZone: "America/Sao_Paulo" });
    const content = m.body?.content?.substring(0, MESSAGE_CONTENT_MAX_LENGTH) ?? "";
    return `- **${from}** (${date})\n  ${content}`;
  });

  return `## Mensagens do Canal (${formatted.length})\n\n${formatted.join("\n\n")}`;
}

export function formatChatList(chats: GraphChat[]): string {
  if (chats.length === 0) return "Nenhum chat encontrado.";

  const formatted = chats.map((c) => {
    const topic = c.topic ?? "Chat sem título";
    const type = c.chatType ?? "unknown";
    const updated = c.lastUpdatedDateTime
      ? new Date(c.lastUpdatedDateTime).toLocaleString("pt-BR", { timeZone: "America/Sao_Paulo" })
      : "N/A";
    const members =
      c.members?.map((m) => m.displayName).join(", ") ?? "";
    return `- **${topic}** (${type})\n  Membros: ${members}\n  Última atualização: ${updated}\n  ID: ${c.id}`;
  });

  return `## Chats (${formatted.length})\n\n${formatted.join("\n\n")}`;
}
