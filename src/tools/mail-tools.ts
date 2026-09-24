import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import type { MailService } from "../services/mail.js";
import { safeTool } from "../utils/safe-tool.js";
import {
  formatEmailList,
  formatEmailSearchResults,
  formatEmailDetail,
  formatMailFolders,
} from "../formatters/mail.js";
import { formatStoredDownloads } from "../formatters/download.js";
import { formatSize } from "../utils/format.js";
import { DEFAULT_EMAIL_BODY_MAX_LENGTH, DOWNLOAD_MAX_BYTES, MAX_EMAIL_BODY_MAX_LENGTH } from "../constants.js";

function resolveBodyLimit(pedido: number): number {
  if (!Number.isFinite(pedido) || pedido <= 0) return DEFAULT_EMAIL_BODY_MAX_LENGTH;
  return Math.min(pedido, MAX_EMAIL_BODY_MAX_LENGTH);
}

export function registerMailTools(server: McpServer, mail: MailService) {
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
      const aliased = emails.map(e => ({ ...e, id: mail.registerAlias(e.id) }));
      return {
        content: [{ type: "text" as const, text: formatEmailList(aliased) }],
      };
    })
  );

  server.tool(
    "search-emails",
    "Busca emails por texto (KQL). Pesquisa no assunto, corpo e remetente. Agrupa por thread (conversationId) e retorna a mensagem mais recente de cada.",
    {
      query: z.string().min(1).describe("Texto para buscar nos emails"),
      top: z.number().optional().describe("Número máximo de threads a retornar (padrão: 10)"),
    },
    safeTool(async (params) => {
      const emails = await mail.searchEmails(params.query, params.top);
      const aliased = emails.map(e => ({ ...e, id: mail.registerAlias(e.id) }));
      return {
        content: [{ type: "text" as const, text: formatEmailSearchResults(params.query, aliased) }],
      };
    })
  );

  server.tool(
    "read-email",
    "Lê o conteúdo completo de um email pelo ID.",
    {
      messageId: z.string().describe("ID do email (aceita alias curto ex: m1)"),
      format: z.enum(["text", "html"]).optional().default("text").describe("Formato do corpo: 'text' (padrão, mais leve) ou 'html'"),
      maxBodyLength: z.number().default(DEFAULT_EMAIL_BODY_MAX_LENGTH).describe("Truncar corpo após N caracteres (padrão: 4000, máximo: 12000)"),
    },
    safeTool(async (params) => {
      const email = await mail.readEmail(mail.resolveId(params.messageId), params.format);
      return {
        content: [{ type: "text" as const, text: formatEmailDetail(email, resolveBodyLimit(params.maxBodyLength)) }],
      };
    })
  );

  server.tool(
    "download-email-attachment",
    `Baixa anexos de um email para o disco local e retorna o caminho de cada arquivo, para anexar na resposta. Sem attachmentName, baixa todos os anexos que não são imagens embutidas no corpo. Email anexado vira arquivo .eml. Limite: ${formatSize(DOWNLOAD_MAX_BYTES)} por arquivo.`,
    {
      messageId: z.string().describe("ID do email (aceita alias curto ex: m1)"),
      attachmentName: z.string().optional().describe("Nome ou parte do nome do anexo, como aparece em read-email"),
    },
    safeTool(async (params) => {
      const downloads = await mail.downloadAttachments(mail.resolveId(params.messageId), params.attachmentName);
      return {
        content: [{ type: "text" as const, text: formatStoredDownloads(downloads) }],
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
      contentType: z.enum(["Text", "HTML"]).optional().describe("Tipo de conteúdo: 'Text' ou 'HTML' (padrão: Text)"),
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
      await mail.replyEmail(mail.resolveId(params.messageId), params.comment);
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
      return {
        content: [{ type: "text" as const, text: formatMailFolders(folders) }],
      };
    })
  );
}
