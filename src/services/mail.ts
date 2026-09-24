import type { AuthProvider } from "../types/auth.js";
import type {
  GraphEmailMessage,
  GraphMailFolder,
  GraphPagedResponse,
} from "../types/graph.js";
import { toRecipient } from "../types/graph.js";
import { graphFetch, graphFetchVoid } from "../utils/graph-client.js";
import { SCOPES, DEFAULT_PAGE_SIZE_SMALL } from "../constants.js";
import { createGetToken } from "../utils/auth-helper.js";
import type { StoredDownload } from "../utils/download-store.js";
import { downloadOutlookAttachments } from "../utils/outlook-attachments.js";

export interface ListEmailsParams {
  folder?: string;
  top?: number;
  skip?: number;
  filter?: string;
}

export interface SendEmailParams {
  to: string[];
  subject: string;
  body: string;
  cc?: string[];
  bcc?: string[];
  contentType?: "Text" | "HTML";
}

const RECEIVED_DATE_TIME_LOWER_BOUND = "receivedDateTime ge 1900-01-01T00:00:00Z";

export function buildOrderedMessageFilter(filter: string): string {
  return `${RECEIVED_DATE_TIME_LOWER_BOUND} and (${filter.trim()})`;
}

function escapeKqlPhrase(text: string): string {
  return text.replace(/\\/g, "\\\\").replace(/"/g, '\\"');
}

function toGraphSearchPhrase(query: string): string {
  const bruta = query.trim();
  const semAspasExternas = bruta.replace(/^"+|"+$/g, "").trim() || bruta;
  return `"\\"${escapeKqlPhrase(semAspasExternas)}\\""`;
}

export function createMailService(auth: AuthProvider) {
  const getToken = createGetToken(auth, SCOPES.MAIL);
  const aliasToId = new Map<string, string>();
  const idToAlias = new Map<string, string>();
  let aliasCounter = 0;

  function registerAlias(realId: string): string {
    const existing = idToAlias.get(realId);
    if (existing) return existing;
    const alias = `m${++aliasCounter}`;
    aliasToId.set(alias, realId);
    idToAlias.set(realId, alias);
    return alias;
  }

  function resolveId(idOrAlias: string): string {
    return aliasToId.get(idOrAlias) ?? idOrAlias;
  }

  async function listEmails(params: ListEmailsParams = {}): Promise<GraphEmailMessage[]> {
    const token = await getToken();
    const { folder, top = DEFAULT_PAGE_SIZE_SMALL, skip = 0, filter } = params;

    const basePath = folder ? `/me/mailFolders/${folder}/messages` : "/me/messages";
    const queryParams = new URLSearchParams({
      $top: String(top),
      $skip: String(skip),
      $select: "id,subject,from,toRecipients,receivedDateTime,isRead,hasAttachments,bodyPreview",
      $orderby: "receivedDateTime desc",
    });
    if (filter?.trim()) queryParams.set("$filter", buildOrderedMessageFilter(filter));

    const result = await graphFetch<GraphPagedResponse<GraphEmailMessage>>(
      token,
      `${basePath}?${queryParams}`
    );
    return result.value;
  }

  async function searchEmails(
    query: string,
    top: number = DEFAULT_PAGE_SIZE_SMALL
  ): Promise<GraphEmailMessage[]> {
    const token = await getToken();
    const queryParams = new URLSearchParams({
      $search: toGraphSearchPhrase(query),
      $top: String(top * 3),
      $select: "id,conversationId,subject,from,toRecipients,receivedDateTime,isRead,hasAttachments,bodyPreview",
    });

    const result = await graphFetch<GraphPagedResponse<GraphEmailMessage>>(
      token,
      `/me/messages?${queryParams}`
    );
    return deduplicateByConversation(result.value, top);
  }

  async function readEmail(
    messageId: string,
    format: "text" | "html" = "text"
  ): Promise<GraphEmailMessage> {
    const token = await getToken();
    return graphFetch<GraphEmailMessage>(
      token,
      `/me/messages/${messageId}?$select=id,subject,from,toRecipients,ccRecipients,receivedDateTime,body,hasAttachments,attachments`,
      { headers: { Prefer: `outlook.body-content-type="${format}"` } }
    );
  }

  async function sendEmail(params: SendEmailParams): Promise<{ success: true }> {
    const token = await getToken();
    const { to, subject, body, cc, bcc, contentType = "Text" } = params;

    const toRecipients = to.map(toRecipient);
    const ccRecipients = cc?.map(toRecipient);
    const bccRecipients = bcc?.map(toRecipient);

    await graphFetchVoid(token, "/me/sendMail", {
      method: "POST",
      body: JSON.stringify({
        message: {
          subject,
          body: {
            contentType,
            content: body,
          },
          toRecipients,
          ccRecipients,
          bccRecipients,
        },
      }),
    });

    return { success: true };
  }

  async function replyEmail(
    messageId: string,
    comment: string
  ): Promise<{ success: true }> {
    const token = await getToken();
    await graphFetchVoid(token, `/me/messages/${messageId}/reply`, {
      method: "POST",
      body: JSON.stringify({ comment }),
    });
    return { success: true };
  }

  async function downloadAttachments(
    messageId: string,
    attachmentName?: string
  ): Promise<StoredDownload[]> {
    const token = await getToken();
    return downloadOutlookAttachments(token, `/me/messages/${messageId}/attachments`, attachmentName);
  }

  async function listMailFolders(): Promise<GraphMailFolder[]> {
    const token = await getToken();
    const result = await graphFetch<GraphPagedResponse<GraphMailFolder>>(
      token,
      "/me/mailFolders?$select=id,displayName,totalItemCount,unreadItemCount"
    );
    return result.value;
  }

  return {
    listEmails, searchEmails, readEmail, sendEmail, replyEmail, listMailFolders,
    downloadAttachments, registerAlias, resolveId,
  };
}

function deduplicateByConversation(
  messages: readonly GraphEmailMessage[],
  limit: number
): GraphEmailMessage[] {
  const byConversation = new Map<string, GraphEmailMessage>();
  for (const msg of messages) {
    const key = msg.conversationId ?? msg.id;
    const existing = byConversation.get(key);
    if (!existing || msg.receivedDateTime > existing.receivedDateTime) {
      byConversation.set(key, msg);
    }
  }
  return [...byConversation.values()]
    .sort((a, b) => b.receivedDateTime.localeCompare(a.receivedDateTime))
    .slice(0, limit);
}

export type MailService = ReturnType<typeof createMailService>;
