import { msalClient } from "../auth/msal-client.ts";
import { graphFetch } from "../utils/errors.ts";

const TEAMS_SCOPES = [
  "Team.ReadBasic.All",
  "Channel.ReadBasic.All",
  "ChannelMessage.Send",
  "Chat.Read",
  "Chat.ReadWrite",
];

async function getToken() {
  return msalClient.getAccessToken(TEAMS_SCOPES);
}

export async function listTeams() {
  const token = await getToken();
  const result = await graphFetch(
    token,
    "/me/joinedTeams?$select=id,displayName,description"
  );
  return result.value;
}

export async function listChannels(teamId: string) {
  const token = await getToken();
  const result = await graphFetch(
    token,
    `/teams/${teamId}/channels?$select=id,displayName,description,membershipType`
  );
  return result.value;
}

export interface ListChannelMessagesParams {
  teamId: string;
  channelId: string;
  top?: number;
}

export async function listChannelMessages(params: ListChannelMessagesParams) {
  const token = await getToken();
  const { teamId, channelId, top = 20 } = params;
  const result = await graphFetch(
    token,
    `/teams/${teamId}/channels/${channelId}/messages?$top=${top}`
  );
  return result.value;
}

export interface SendChannelMessageParams {
  teamId: string;
  channelId: string;
  content: string;
  contentType?: "text" | "html";
}

export async function sendChannelMessage(params: SendChannelMessageParams) {
  const token = await getToken();
  const { teamId, channelId, content, contentType = "text" } = params;

  const result = await graphFetch(
    token,
    `/teams/${teamId}/channels/${channelId}/messages`,
    {
      method: "POST",
      body: JSON.stringify({
        body: { contentType, content },
      }),
    }
  );
  return result;
}

export async function listChats(top: number = 20) {
  const token = await getToken();
  const result = await graphFetch(
    token,
    `/me/chats?$top=${top}&$select=id,topic,chatType,lastUpdatedDateTime&$expand=members($select=displayName,email)`
  );
  return result.value;
}

export interface SendChatMessageParams {
  chatId: string;
  content: string;
  contentType?: "text" | "html";
}

export async function sendChatMessage(params: SendChatMessageParams) {
  const token = await getToken();
  const { chatId, content, contentType = "text" } = params;

  const result = await graphFetch(token, `/me/chats/${chatId}/messages`, {
    method: "POST",
    body: JSON.stringify({
      body: { contentType, content },
    }),
  });
  return result;
}
