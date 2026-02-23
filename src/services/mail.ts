import { msalClient } from "../auth/msal-client.js";
import { graphFetch } from "../utils/errors.js";

const MAIL_SCOPES = ["Mail.Read", "Mail.Send", "Mail.ReadWrite"];

async function getToken() {
  return msalClient.getAccessToken(MAIL_SCOPES);
}

export interface ListEmailsParams {
  folder?: string;
  top?: number;
  skip?: number;
  filter?: string;
}

export async function listEmails(params: ListEmailsParams = {}) {
  const token = await getToken();
  const { folder, top = 10, skip = 0, filter } = params;

  const basePath = folder ? `/me/mailFolders/${folder}/messages` : "/me/messages";
  const queryParams = new URLSearchParams({
    $top: String(top),
    $skip: String(skip),
    $select: "id,subject,from,toRecipients,receivedDateTime,isRead,hasAttachments,bodyPreview",
    $orderby: "receivedDateTime desc",
  });
  if (filter) queryParams.set("$filter", filter);

  const result = await graphFetch(token, `${basePath}?${queryParams}`);
  return result.value;
}

export async function searchEmails(query: string, top: number = 10) {
  const token = await getToken();
  const queryParams = new URLSearchParams({
    $search: `"${query}"`,
    $top: String(top),
    $select: "id,subject,from,toRecipients,receivedDateTime,isRead,hasAttachments,bodyPreview",
  });

  const result = await graphFetch(token, `/me/messages?${queryParams}`);
  return result.value;
}

export async function readEmail(messageId: string) {
  const token = await getToken();
  const result = await graphFetch(
    token,
    `/me/messages/${messageId}?$select=id,subject,from,toRecipients,ccRecipients,receivedDateTime,body,hasAttachments,attachments`
  );
  return result;
}

export interface SendEmailParams {
  to: string[];
  subject: string;
  body: string;
  cc?: string[];
  bcc?: string[];
  isHtml?: boolean;
}

export async function sendEmail(params: SendEmailParams) {
  const token = await getToken();
  const { to, subject, body, cc, bcc, isHtml = false } = params;

  const toRecipients = to.map((email) => ({
    emailAddress: { address: email },
  }));
  const ccRecipients = cc?.map((email) => ({
    emailAddress: { address: email },
  }));
  const bccRecipients = bcc?.map((email) => ({
    emailAddress: { address: email },
  }));

  await graphFetch(token, "/me/sendMail", {
    method: "POST",
    body: JSON.stringify({
      message: {
        subject,
        body: {
          contentType: isHtml ? "HTML" : "Text",
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

export async function replyEmail(messageId: string, comment: string) {
  const token = await getToken();
  await graphFetch(token, `/me/messages/${messageId}/reply`, {
    method: "POST",
    body: JSON.stringify({ comment }),
  });
  return { success: true };
}

export async function listMailFolders() {
  const token = await getToken();
  const result = await graphFetch(
    token,
    "/me/mailFolders?$select=id,displayName,totalItemCount,unreadItemCount"
  );
  return result.value;
}
