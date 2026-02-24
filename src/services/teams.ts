import type { AuthProvider } from "../types/auth.js";
import type {
  GraphTeam,
  GraphChannel,
  GraphChatMessage,
  GraphChat,
  GraphPagedResponse,
} from "../types/graph.js";
import { graphFetch } from "../utils/graph-client.js";
import { SCOPES, DEFAULT_PAGE_SIZE_LARGE } from "../constants.js";
import { createGetToken } from "../utils/auth-helper.js";

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

export function createTeamsService(auth: AuthProvider) {
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
      `/me/chats?$top=${top}&$select=id,topic,chatType,lastUpdatedDateTime&$expand=members($select=displayName,email)`
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

  return {
    listTeams,
    listChannels,
    listChannelMessages,
    sendChannelMessage,
    listChats,
    sendChatMessage,
  };
}

export type TeamsService = ReturnType<typeof createTeamsService>;
