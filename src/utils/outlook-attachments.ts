import type { GraphMessageAttachment, GraphPagedResponse } from "../types/graph.js";
import { graphFetch, graphFetchBinary } from "./graph-client.js";
import { assertDownloadable, storeDownload, type StoredDownload } from "./download-store.js";

export const ITEM_ATTACHMENT = "#microsoft.graph.itemAttachment";
export const REFERENCE_ATTACHMENT = "#microsoft.graph.referenceAttachment";

export function listAttachmentNames(attachments: readonly GraphMessageAttachment[]): string {
  return attachments.length > 0 ? attachments.map((a) => `'${a.name}'`).join(", ") : "nenhum";
}

export function findAttachmentsByName(
  attachments: readonly GraphMessageAttachment[],
  attachmentName: string
): GraphMessageAttachment[] {
  const wanted = attachmentName.trim().toLowerCase();
  const exactMatches = attachments.filter((a) => a.name.toLowerCase() === wanted);
  if (exactMatches.length > 0) return exactMatches;
  return attachments.filter((a) => a.name.toLowerCase().includes(wanted));
}

function referenceAttachmentHint(links: readonly GraphMessageAttachment[]): string {
  if (links.length === 0) return "";
  return ` Links do OneDrive/SharePoint não são baixáveis como anexo — copie o link do corpo do item e use download-drive-file com shareUrl: ${listAttachmentNames(links)}.`;
}

export function describeMissingAttachment(
  attachmentName: string,
  downloadable: readonly GraphMessageAttachment[],
  links: readonly GraphMessageAttachment[]
): string {
  return `Nenhum anexo corresponde a '${attachmentName}'. Anexos disponíveis: ${listAttachmentNames(downloadable)}.${referenceAttachmentHint(links)}`;
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
  throw new Error(`O item não tem anexos para baixar.${referenceAttachmentHint(links)}`);
}

export function attachmentFileName(attachment: GraphMessageAttachment): string {
  const isAttachedEmail = attachment["@odata.type"] === ITEM_ATTACHMENT;
  return isAttachedEmail && !/\.eml$/i.test(attachment.name) ? `${attachment.name}.eml` : attachment.name;
}

export async function downloadOutlookAttachments(
  token: string,
  attachmentsUrl: string,
  attachmentName?: string
): Promise<StoredDownload[]> {
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
