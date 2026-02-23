import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import * as mail from "../services/mail.ts";
import { safeTool } from "../utils/errors.ts";

export function registerMailTools(server: McpServer) {
  server.tool(
    "list-emails",
    "Lista emails do Outlook. Por padrão retorna os 10 emails mais recentes da Inbox.",
    {
      folder: z.string().optional().describe("ID da pasta (ex: 'Inbox', 'SentItems', 'Drafts')"),
      top: z.number().optional().describe("Número de emails a retornar (padrão: 10)"),
      skip: z.number().optional().describe("Número de emails a pular (para paginação)"),
      filter: z.string().optional().describe("Filtro OData (ex: \"isRead eq false\")"),
    },
    safeTool(async (params) => {
      const emails = await mail.listEmails(params);
      const formatted = emails.map((e: any) => {
        const from = e.from?.emailAddress?.address ?? "unknown";
        const date = new Date(e.receivedDateTime).toLocaleString("pt-BR");
        const read = e.isRead ? "" : " [NÃO LIDO]";
        const attach = e.hasAttachments ? " [ANEXO]" : "";
        return `- **${e.subject}**${read}${attach}\n  De: ${from} | ${date}\n  ID: ${e.id}\n  ${e.bodyPreview?.substring(0, 100) ?? ""}...`;
      });

      return {
        content: [
          {
            type: "text" as const,
            text: formatted.length > 0
              ? `## Emails (${formatted.length})\n\n${formatted.join("\n\n")}`
              : "Nenhum email encontrado.",
          },
        ],
      };
    })
  );

  server.tool(
    "search-emails",
    "Busca emails por texto (KQL). Pesquisa no assunto, corpo e remetente.",
    {
      query: z.string().describe("Texto para buscar nos emails"),
      top: z.number().optional().describe("Número máximo de resultados (padrão: 10)"),
    },
    safeTool(async (params) => {
      const emails = await mail.searchEmails(params.query, params.top);
      const formatted = emails.map((e: any) => {
        const from = e.from?.emailAddress?.address ?? "unknown";
        const date = new Date(e.receivedDateTime).toLocaleString("pt-BR");
        return `- **${e.subject}**\n  De: ${from} | ${date}\n  ID: ${e.id}`;
      });

      return {
        content: [
          {
            type: "text" as const,
            text: formatted.length > 0
              ? `## Resultados da busca "${params.query}" (${formatted.length})\n\n${formatted.join("\n\n")}`
              : `Nenhum email encontrado para "${params.query}".`,
          },
        ],
      };
    })
  );

  server.tool(
    "read-email",
    "Lê o conteúdo completo de um email pelo ID.",
    {
      messageId: z.string().describe("ID do email a ser lido"),
    },
    safeTool(async (params) => {
      const email = await mail.readEmail(params.messageId);
      const from = email.from?.emailAddress?.address ?? "unknown";
      const to = email.toRecipients?.map((r: any) => r.emailAddress?.address).join(", ") ?? "";
      const cc = email.ccRecipients?.map((r: any) => r.emailAddress?.address).join(", ");
      const date = new Date(email.receivedDateTime).toLocaleString("pt-BR");
      const bodyContent = email.body?.content ?? "(sem conteúdo)";

      let text = [
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

      return {
        content: [{ type: "text" as const, text: text.join("\n") }],
      };
    })
  );

  server.tool(
    "send-email",
    "Envia um novo email.",
    {
      to: z.array(z.string()).describe("Lista de endereços de email dos destinatários"),
      subject: z.string().describe("Assunto do email"),
      body: z.string().describe("Corpo do email"),
      cc: z.array(z.string()).optional().describe("Lista de endereços CC"),
      bcc: z.array(z.string()).optional().describe("Lista de endereços BCC"),
      isHtml: z.boolean().optional().describe("Se true, o corpo é HTML (padrão: texto plano)"),
    },
    safeTool(async (params) => {
      await mail.sendEmail(params);
      return {
        content: [
          {
            type: "text" as const,
            text: `Email enviado com sucesso para: ${params.to.join(", ")}`,
          },
        ],
      };
    })
  );

  server.tool(
    "reply-email",
    "Responde a um email existente.",
    {
      messageId: z.string().describe("ID do email a responder"),
      comment: z.string().describe("Texto da resposta"),
    },
    safeTool(async (params) => {
      await mail.replyEmail(params.messageId, params.comment);
      return {
        content: [
          { type: "text" as const, text: "Resposta enviada com sucesso." },
        ],
      };
    })
  );

  server.tool(
    "list-mail-folders",
    "Lista as pastas de email (Inbox, Sent Items, etc.).",
    {},
    safeTool(async () => {
      const folders = await mail.listMailFolders();
      const formatted = folders.map(
        (f: any) =>
          `- **${f.displayName}** — ${f.unreadItemCount} não lidos / ${f.totalItemCount} total\n  ID: ${f.id}`
      );

      return {
        content: [
          {
            type: "text" as const,
            text: `## Pastas de Email\n\n${formatted.join("\n\n")}`,
          },
        ],
      };
    })
  );
}
