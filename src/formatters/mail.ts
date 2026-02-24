import type { GraphEmailMessage, GraphMailFolder } from "../types/graph.js";
import { recipientAddress, recipientAddresses } from "../types/graph.js";
import { BODY_PREVIEW_MAX_LENGTH } from "../constants.js";
import { formatDateBR } from "../utils/date.js";

export function formatEmailList(emails: GraphEmailMessage[]): string {
  if (emails.length === 0) return "Nenhum email encontrado.";

  const formatted = emails.map((e) => {
    const from = recipientAddress(e.from);
    const date = formatDateBR(e.receivedDateTime);
    const read = e.isRead ? "" : " [NÃO LIDO]";
    const attach = e.hasAttachments ? " [ANEXO]" : "";
    return `- **${e.subject}**${read}${attach}\n  De: ${from} | ${date}\n  ID: ${e.id}\n  ${e.bodyPreview?.substring(0, BODY_PREVIEW_MAX_LENGTH) ?? ""}...`;
  });

  return `## Emails (${formatted.length})\n\n${formatted.join("\n\n")}`;
}

export function formatEmailSearchResults(
  query: string,
  emails: GraphEmailMessage[]
): string {
  if (emails.length === 0) return `Nenhum email encontrado para "${query}".`;

  const formatted = emails.map((e) => {
    const from = recipientAddress(e.from);
    const date = formatDateBR(e.receivedDateTime);
    return `- **${e.subject}**\n  De: ${from} | ${date}\n  ID: ${e.id}`;
  });

  return `## Resultados da busca "${query}" (${formatted.length})\n\n${formatted.join("\n\n")}`;
}

export function formatEmailDetail(email: GraphEmailMessage): string {
  const from = recipientAddress(email.from);
  const to = recipientAddresses(email.toRecipients);
  const cc = recipientAddresses(email.ccRecipients);
  const date = formatDateBR(email.receivedDateTime);
  const bodyContent = email.body?.content ?? "(sem conteúdo)";

  const text = [
    `## ${email.subject}`,
    "",
    `**De:** ${from}`,
    `**Para:** ${to}`,
  ];
  if (cc) text.push(`**CC:** ${cc}`);
  text.push(`**Data:** ${date}`);
  text.push("");
  text.push(bodyContent);

  if (email.hasAttachments && email.attachments?.length) {
    text.push("");
    text.push("### Anexos");
    for (const att of email.attachments) {
      text.push(`- ${att.name} (${att.contentType}, ${att.size} bytes)`);
    }
  }

  return text.join("\n");
}

export function formatMailFolders(folders: GraphMailFolder[]): string {
  const formatted = folders.map(
    (f) =>
      `- **${f.displayName}** — ${f.unreadItemCount} não lidos / ${f.totalItemCount} total\n  ID: ${f.id}`
  );

  return `## Pastas de Email\n\n${formatted.join("\n\n")}`;
}
