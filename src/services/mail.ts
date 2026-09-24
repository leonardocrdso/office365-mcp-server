import type { AuthProvider } from "../types/auth.js";
import type {
  GraphEmailMessage,
  GraphMailFolder,
  GraphMessageAttachment,
  GraphPagedResponse,
} from "../types/graph.js";
import { toRecipient } from "../types/graph.js";
import { graphFetch, graphFetchBinary, graphFetchVoid } from "../utils/graph-client.js";
import { SCOPES, DEFAULT_PAGE_SIZE_SMALL } from "../constants.js";
import { createGetToken } from "../utils/auth-helper.js";
import { assertDownloadable, storeDownload, type StoredDownload } from "../utils/download-store.js";

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

const ITEM_ATTACHMENT = "#microsoft.graph.itemAttachment";
const REFERENCE_ATTACHMENT = "#microsoft.graph.referenceAttachment";

function listAttachmentNames(attachments: readonly GraphMessageAttachment[]): string {
  return attachments.length > 0 ? attachments.map((a) => `'${a.name}'`).join(", ") : "nenhum";
}

function findAttachmentsByName(
  attachments: readonly GraphMessageAttachment[],
  attachmentName: string
): GraphMessageAttachment[] {
  const wanted = attachmentName.trim().toLowerCase();
  const exactMatches = attachments.filter((a) => a.name.toLowerCase() === wanted);
  if (exactMatches.length > 0) return exactMatches;
  return attachments.filter((a) => a.name.toLowerCase().includes(wanted));
}

function describeMissingAttachment(
  attachmentName: string,
  downloadable: readonly GraphMessageAttachment[],
  links: readonly GraphMessageAttachment[]
): string {
  const linkHint = links.length > 0
    ? ` Links do OneDrive/SharePoint no email (use resolve-share-link e download-drive-file): ${listAttachmentNames(links)}.`
    : "";
  return `Nenhum anexo corresponde a '${attachmentName}'. Anexos disponíveis: ${listAttachmentNames(downloadable)}.${linkHint}`;
}

export function selectAttachments(
  attachments: readonly GraphMessageAttachment[],
  attachmentName?: string
): GraphMessageAttachment[] {
  const links = attachments.filter((a) => a["@odata.type"] === REFERENCE_ATTACHMENT);
  const downloadable = attachments.filter((a) => a["@odata.type"] !== REFERENCE_ATTACHMENT);
  if (attachmentName) {
    const matches = findAttachmentsByName(downloadable, attachmentName);
    if (matches.length > 0) return matches;
    throw new Error(describeMissingAttachment(attachmentName, downloadable, links));
  }
  const visible = downloadable.filter((a) => !a.isInline);
  if (visible.length > 0) return visible;
  throw new Error(`O email não tem anexos para baixar.${links.length > 0 ? ` Só links: ${listAttachmentNames(links)}.` : ""}`);
}

export function attachmentFileName(attachment: GraphMessageAttachment): string {
  const isAttachedEmail = attachment["@odata.type"] === ITEM_ATTACHMENT;
  return isAttachedEmail && !/\.eml$/i.test(attachment.name) ? `${attachment.name}.eml` : attachment.name;
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
    if (filter) queryParams.set("$filter", filter);

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
    const attachmentsUrl = `/me/messages/${messageId}/attachments`;
    const listing = await graphFetch<GraphPagedResponse<GraphMessageAttachment>>(
      token,
      `${attachmentsUrl}?$select=id,name,contentType,size,isInline`
    );
    const selected = selectAttachments(listing.value, attachmentName);
    selected.forEach((attachment) => assertDownloadable(attachment.name, attachment.size));
    const downloads: StoredDownload[] = [];
    for (const attachment of selected) {
      const content = await graphFetchBinary(token, `${attachmentsUrl}/${encodeURIComponent(attachment.id)}/$value`);
      downloads.push(await storeDownload(attachmentFileName(attachment), content));
    }
    return downloads;
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
