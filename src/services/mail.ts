import type { AuthProvider } from "../types/auth.js";
import type {
  GraphEmailMessage,
  GraphMailFolder,
  GraphPagedResponse,
} from "../types/graph.js";
import { graphFetch, graphFetchVoid } from "../utils/graph-client.js";
import { SCOPES, DEFAULT_PAGE_SIZE_SMALL } from "../constants.js";

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

export function createMailService(auth: AuthProvider) {
  async function getToken() {
    return auth.getAccessToken([...SCOPES.MAIL]);
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
      $search: `"${query}"`,
      $top: String(top),
      $select: "id,subject,from,toRecipients,receivedDateTime,isRead,hasAttachments,bodyPreview",
    });

    const result = await graphFetch<GraphPagedResponse<GraphEmailMessage>>(
      token,
      `/me/messages?${queryParams}`
    );
    return result.value;
  }

  async function readEmail(messageId: string): Promise<GraphEmailMessage> {
    const token = await getToken();
    return graphFetch<GraphEmailMessage>(
      token,
      `/me/messages/${messageId}?$select=id,subject,from,toRecipients,ccRecipients,receivedDateTime,body,hasAttachments,attachments`
    );
  }

  async function sendEmail(params: SendEmailParams): Promise<{ success: true }> {
    const token = await getToken();
    const { to, subject, body, cc, bcc, contentType = "Text" } = params;

    const toRecipients = to.map((email) => ({
      emailAddress: { address: email },
    }));
    const ccRecipients = cc?.map((email) => ({
      emailAddress: { address: email },
    }));
    const bccRecipients = bcc?.map((email) => ({
      emailAddress: { address: email },
    }));

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

  async function listMailFolders(): Promise<GraphMailFolder[]> {
    const token = await getToken();
    const result = await graphFetch<GraphPagedResponse<GraphMailFolder>>(
      token,
      "/me/mailFolders?$select=id,displayName,totalItemCount,unreadItemCount"
    );
    return result.value;
  }

  return { listEmails, searchEmails, readEmail, sendEmail, replyEmail, listMailFolders };
}

export type MailService = ReturnType<typeof createMailService>;
