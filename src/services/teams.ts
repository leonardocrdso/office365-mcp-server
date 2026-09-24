import type { AuthProvider } from "../types/auth.js";
import type {
  GraphTeam,
  GraphChannel,
  GraphChatMessage,
  GraphChatMessageAttachment,
  GraphChat,
  GraphPagedResponse,
} from "../types/graph.js";
import { graphFetch, graphFetchBinary } from "../utils/graph-client.js";
import { SCOPES, DEFAULT_PAGE_SIZE_LARGE } from "../constants.js";
import { createGetToken } from "../utils/auth-helper.js";
import { storeDownload, type StoredDownload } from "../utils/download-store.js";

export interface ListChannelMessagesParams {
  teamId: string;
  channelId: string;
  top?: number;
}

export interface SendChannelMessageParams {
  teamId: string;
  channelId: string;
  content: string;
  contentType?: "text" | "html";
}

export interface SendChatMessageParams {
  chatId: string;
  content: string;
  contentType?: "text" | "html";
}

export interface ListChatMessagesParams {
  chatId: string;
  top?: number;
}

export interface ListChannelMessageRepliesParams {
  teamId: string;
  channelId: string;
  messageId: string;
  top?: number;
}

export type TeamsMessageLocation =
  | { kind: "chat"; chatId: string; messageId: string }
  | { kind: "channel"; teamId: string; channelId: string; messageId: string }
  | { kind: "channelReply"; teamId: string; channelId: string; parentMessageId: string; messageId: string };

export interface TeamsServiceDeps {
  downloadSharedFile: (shareUrl: string) => Promise<StoredDownload>;
}

const HOSTED_CONTENT_REGEX = /hostedContents\/([^/"]+)\/\$value/g;

function channelMessagesPath(teamId: string, channelId: string): string {
  return `/teams/${teamId}/channels/${channelId}/messages`;
}

function messageBasePath(location: TeamsMessageLocation): string {
  switch (location.kind) {
    case "chat":
      return `/chats/${location.chatId}/messages/${location.messageId}`;
    case "channel":
      return `${channelMessagesPath(location.teamId, location.channelId)}/${location.messageId}`;
    case "channelReply":
      return `${channelMessagesPath(location.teamId, location.channelId)}/${location.parentMessageId}/replies/${location.messageId}`;
  }
}

export function extractHostedContentIds(bodyHtml: string): readonly string[] {
  return [...bodyHtml.matchAll(HOSTED_CONTENT_REGEX)].map((match) => match[1] ?? "").filter(Boolean);
}

function isDownloadableReference(
  attachment: GraphChatMessageAttachment
): attachment is GraphChatMessageAttachment & { contentUrl: string } {
  return attachment.contentType === "reference" && Boolean(attachment.contentUrl);
}

function matchByName<T extends { name?: string }>(items: readonly T[], name: string): T[] {
  const wanted = name.trim().toLowerCase();
  const named = items.filter((item) => Boolean(item.name));
  const exact = named.filter((item) => (item.name ?? "").toLowerCase() === wanted);
  if (exact.length > 0) return exact;
  return named.filter((item) => (item.name ?? "").toLowerCase().includes(wanted));
}

function describeAttachments(attachments: readonly { name?: string }[]): string {
  const names = attachments.map((a) => a.name).filter((name): name is string => Boolean(name));
  return names.length > 0 ? names.map((n) => `'${n}'`).join(", ") : "nenhum";
}

function noFilesToDownloadError(allAttachments: readonly GraphChatMessageAttachment[]): Error {
  const names = describeAttachments(allAttachments);
  if (names === "nenhum") return new Error("A mensagem não tem arquivos para baixar.");
  return new Error(`A mensagem não tem arquivos para baixar. Anexos disponíveis: ${names}.`);
}

function selectReferenceAttachments(
  allAttachments: readonly GraphChatMessageAttachment[],
  attachmentName?: string
): readonly (GraphChatMessageAttachment & { contentUrl: string })[] {
  const referenceAttachments = allAttachments.filter(isDownloadableReference);
  if (!attachmentName) return referenceAttachments;
  const matches = matchByName(referenceAttachments, attachmentName);
  if (matches.length > 0) return matches;
  throw new Error(
    `Nenhum anexo corresponde a '${attachmentName}'. Anexos disponíveis: ${describeAttachments(allAttachments)}.`
  );
}

function extensionFromMagicBytes(bytes: Uint8Array): string {
  if (bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47) return "png";
  if (bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff) return "jpg";
  if (bytes[0] === 0x47 && bytes[1] === 0x49 && bytes[2] === 0x46 && bytes[3] === 0x38) return "gif";
  return "bin";
}

export function createTeamsService(auth: AuthProvider, deps: TeamsServiceDeps) {
  const getToken = createGetToken(auth, SCOPES.TEAMS);

  async function listTeams(): Promise<GraphTeam[]> {
    const token = await getToken();
    const result = await graphFetch<GraphPagedResponse<GraphTeam>>(
      token,
      "/me/joinedTeams?$select=id,displayName,description"
    );
    return result.value;
  }

  async function listChannels(teamId: string): Promise<GraphChannel[]> {
    const token = await getToken();
    const result = await graphFetch<GraphPagedResponse<GraphChannel>>(
      token,
      `/teams/${teamId}/channels?$select=id,displayName,description,membershipType`
    );
    return result.value;
  }

  async function listChannelMessages(
    params: ListChannelMessagesParams
  ): Promise<GraphChatMessage[]> {
    const token = await getToken();
    const { teamId, channelId, top = DEFAULT_PAGE_SIZE_LARGE } = params;
    const result = await graphFetch<GraphPagedResponse<GraphChatMessage>>(
      token,
      `/teams/${teamId}/channels/${channelId}/messages?$top=${top}`
    );
    return result.value;
  }

  async function listChannelMessageReplies(
    params: ListChannelMessageRepliesParams
  ): Promise<GraphChatMessage[]> {
    const token = await getToken();
    const { teamId, channelId, messageId, top = DEFAULT_PAGE_SIZE_LARGE } = params;
    const result = await graphFetch<GraphPagedResponse<GraphChatMessage>>(
      token,
      `${channelMessagesPath(teamId, channelId)}/${messageId}/replies?$top=${top}`
    );
    return result.value;
  }

  async function sendChannelMessage(
    params: SendChannelMessageParams
  ): Promise<GraphChatMessage> {
    const token = await getToken();
    const { teamId, channelId, content, contentType = "text" } = params;

    return graphFetch<GraphChatMessage>(
      token,
      `/teams/${teamId}/channels/${channelId}/messages`,
      {
        method: "POST",
        body: JSON.stringify({
          body: { contentType, content },
        }),
      }
    );
  }

  async function listChats(top: number = DEFAULT_PAGE_SIZE_LARGE): Promise<GraphChat[]> {
    const token = await getToken();
    const result = await graphFetch<GraphPagedResponse<GraphChat>>(
      token,
      `/me/chats?$top=${top}&$select=id,topic,chatType,lastUpdatedDateTime&$expand=members`
    );
    return result.value;
  }

  async function listChatMessages(
    params: ListChatMessagesParams
  ): Promise<GraphChatMessage[]> {
    const token = await getToken();
    const { chatId, top = DEFAULT_PAGE_SIZE_LARGE } = params;
    const result = await graphFetch<GraphPagedResponse<GraphChatMessage>>(
      token,
      `/me/chats/${chatId}/messages?$top=${top}`
    );
    return result.value;
  }

  async function sendChatMessage(params: SendChatMessageParams): Promise<GraphChatMessage> {
    const token = await getToken();
    const { chatId, content, contentType = "text" } = params;

    return graphFetch<GraphChatMessage>(token, `/me/chats/${chatId}/messages`, {
      method: "POST",
      body: JSON.stringify({
        body: { contentType, content },
      }),
    });
  }

  async function downloadEmbeddedImage(
    location: TeamsMessageLocation,
    hostedContentId: string,
    index: number
  ): Promise<StoredDownload> {
    const token = await getToken();
    const bytes = await graphFetchBinary(
      token,
      `${messageBasePath(location)}/hostedContents/${hostedContentId}/$value`
    );
    return storeDownload(`imagem-${index + 1}.${extensionFromMagicBytes(bytes)}`, bytes);
  }

  async function downloadMessageFiles(
    location: TeamsMessageLocation,
    attachmentName?: string
  ): Promise<StoredDownload[]> {
    const token = await getToken();
    const message = await graphFetch<GraphChatMessage>(token, messageBasePath(location));
    const allAttachments = message.attachments ?? [];
    const selectedAttachments = selectReferenceAttachments(allAttachments, attachmentName);
    const hostedContentIds = attachmentName ? [] : extractHostedContentIds(message.body?.content ?? "");
    if (selectedAttachments.length === 0 && hostedContentIds.length === 0) {
      throw noFilesToDownloadError(allAttachments);
    }

    const downloads: StoredDownload[] = [];
    for (const attachment of selectedAttachments) {
      downloads.push(await deps.downloadSharedFile(attachment.contentUrl));
    }
    for (const [index, hostedContentId] of hostedContentIds.entries()) {
      downloads.push(await downloadEmbeddedImage(location, hostedContentId, index));
    }
    return downloads;
  }

  return {
    listTeams,
    listChannels,
    listChannelMessages,
    listChannelMessageReplies,
    sendChannelMessage,
    listChats,
    listChatMessages,
    sendChatMessage,
    downloadMessageFiles,
  };
}

export type TeamsService = ReturnType<typeof createTeamsService>;
