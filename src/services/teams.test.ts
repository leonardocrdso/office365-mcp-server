import { afterEach, beforeEach, describe, expect, it, mock } from "bun:test";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import type { AuthProvider } from "../types/auth.js";
import type { GraphChatMessage } from "../types/graph.js";
import { createTeamsService, type TeamsMessageLocation } from "./teams.js";

const fakeAuth: AuthProvider = {
  async getAccessToken() {
    return "fake-token";
  },
};

function jsonResponse(body: unknown): Response {
  return new Response(JSON.stringify(body), {
    status: 200,
    headers: { "content-type": "application/json" },
  });
}

function chatMessageWithFilesAndImage(): GraphChatMessage {
  return {
    id: "msg-1",
    createdDateTime: "2026-01-01T00:00:00Z",
    from: { user: { displayName: "Fulano" } },
    body: {
      contentType: "html",
      content:
        '<div>olha essa imagem <img src="https://graph.microsoft.com/v1.0/chats/chat-1/messages/msg-1/hostedContents/abc123/$value"></div>',
    },
    attachments: [
      {
        id: "att-1",
        contentType: "reference",
        contentUrl: "https://contoso.sharepoint.com/planilha.xlsx",
        name: "planilha.xlsx",
      },
      {
        id: "card-1",
        contentType: "application/vnd.microsoft.card.adaptive",
      },
    ],
  };
}

describe("downloadMessageFiles - mensagem de chat", () => {
  const originalFetch = global.fetch;
  let originalHome: string | undefined;
  let tempHome: string;
  let requestedUrls: string[];

  beforeEach(async () => {
    originalHome = process.env.OFFICE365_MCP_HOME;
    tempHome = await mkdtemp(join(tmpdir(), "o365-test-"));
    process.env.OFFICE365_MCP_HOME = tempHome;
    requestedUrls = [];
  });

  afterEach(async () => {
    global.fetch = originalFetch;
    if (originalHome === undefined) delete process.env.OFFICE365_MCP_HOME;
    else process.env.OFFICE365_MCP_HOME = originalHome;
    await rm(tempHome, { recursive: true, force: true });
  });

  const chatLocation: TeamsMessageLocation = { kind: "chat", chatId: "chat-1", messageId: "msg-1" };
  const pngBytes = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 1, 2, 3]);

  function mockMessageThenImage(message: GraphChatMessage): void {
    global.fetch = mock(async (input: RequestInfo | URL) => {
      const url = input.toString();
      requestedUrls.push(url);
      if (url.includes("/hostedContents/")) {
        return new Response(pngBytes, { status: 200, headers: { "content-type": "application/octet-stream" } });
      }
      return jsonResponse(message);
    }) as unknown as typeof fetch;
  }

  it("baixa so o anexo de referencia via downloadSharedFile injetado e a imagem embutida como imagem-1.png", async () => {
    mockMessageThenImage(chatMessageWithFilesAndImage());
    const downloadSharedFile = mock(async (shareUrl: string) => ({
      path: `/tmp/${shareUrl}`,
      fileName: "planilha.xlsx",
      displayName: "planilha.xlsx",
      sizeBytes: 10,
    }));

    const teams = createTeamsService(fakeAuth, { downloadSharedFile });
    const downloads = await teams.downloadMessageFiles(chatLocation);

    expect(downloadSharedFile.mock.calls.length).toBe(1);
    expect(downloadSharedFile.mock.calls[0]?.[0]).toBe("https://contoso.sharepoint.com/planilha.xlsx");

    expect(downloads.length).toBe(2);
    const imagem = downloads.find((d) => d.fileName === "imagem-1.png");
    expect(imagem).toBeDefined();
    if (!imagem) return;
    const written = await readFile(imagem.path);
    expect(new Uint8Array(written)).toEqual(pngBytes);

    expect(requestedUrls[0]).toBe("https://graph.microsoft.com/v1.0/chats/chat-1/messages/msg-1");
    expect(requestedUrls[1]).toContain("/chats/chat-1/messages/msg-1/hostedContents/abc123/$value");
  });

  it("com attachmentName informado, baixa so o anexo correspondente e ignora imagens embutidas", async () => {
    mockMessageThenImage(chatMessageWithFilesAndImage());
    const downloadSharedFile = mock(async (shareUrl: string) => ({
      path: `/tmp/${shareUrl}`,
      fileName: "planilha.xlsx",
      displayName: "planilha.xlsx",
      sizeBytes: 10,
    }));

    const teams = createTeamsService(fakeAuth, { downloadSharedFile });
    const downloads = await teams.downloadMessageFiles(chatLocation, "planilha");

    expect(downloads.length).toBe(1);
    expect(downloadSharedFile.mock.calls.length).toBe(1);
    expect(requestedUrls.some((url) => url.includes("/hostedContents/"))).toBe(false);
  });

  it("lanca erro em pt-BR quando a mensagem nao tem arquivos para baixar", async () => {
    mockMessageThenImage({
      id: "msg-2",
      createdDateTime: "2026-01-01T00:00:00Z",
      body: { contentType: "text", content: "so texto, sem arquivo" },
    });
    const downloadSharedFile = mock(async () => {
      throw new Error("nao deveria ser chamado");
    });

    const teams = createTeamsService(fakeAuth, { downloadSharedFile });

    await expect(teams.downloadMessageFiles(chatLocation)).rejects.toThrow(
      "A mensagem não tem arquivos para baixar."
    );
  });
});

describe("downloadMessageFiles - mensagem de canal", () => {
  const originalFetch = global.fetch;
  let originalHome: string | undefined;
  let tempHome: string;
  let requestedUrls: string[];

  beforeEach(async () => {
    originalHome = process.env.OFFICE365_MCP_HOME;
    tempHome = await mkdtemp(join(tmpdir(), "o365-test-"));
    process.env.OFFICE365_MCP_HOME = tempHome;
    requestedUrls = [];
  });

  afterEach(async () => {
    global.fetch = originalFetch;
    if (originalHome === undefined) delete process.env.OFFICE365_MCP_HOME;
    else process.env.OFFICE365_MCP_HOME = originalHome;
    await rm(tempHome, { recursive: true, force: true });
  });

  it("usa o path /teams/.../channels/.../messages/<id> para localizar a mensagem e a imagem embutida", async () => {
    const pngBytes = new Uint8Array([0x89, 0x50, 0x4e, 0x47]);
    const message: GraphChatMessage = {
      id: "msg-9",
      createdDateTime: "2026-01-01T00:00:00Z",
      body: {
        contentType: "html",
        content:
          '<img src="https://graph.microsoft.com/v1.0/teams/team-1/channels/channel-1/messages/msg-9/hostedContents/xyz/$value">',
      },
    };
    global.fetch = mock(async (input: RequestInfo | URL) => {
      const url = input.toString();
      requestedUrls.push(url);
      if (url.includes("/hostedContents/")) {
        return new Response(pngBytes, { status: 200, headers: { "content-type": "application/octet-stream" } });
      }
      return jsonResponse(message);
    }) as unknown as typeof fetch;

    const downloadSharedFile = mock(async () => {
      throw new Error("nao deveria ser chamado");
    });
    const teams = createTeamsService(fakeAuth, { downloadSharedFile });
    const location: TeamsMessageLocation = { kind: "channel", teamId: "team-1", channelId: "channel-1", messageId: "msg-9" };

    const downloads = await teams.downloadMessageFiles(location);

    expect(downloads.length).toBe(1);
    expect(requestedUrls[0]).toBe("https://graph.microsoft.com/v1.0/teams/team-1/channels/channel-1/messages/msg-9");
    expect(requestedUrls[1]).toContain("/teams/team-1/channels/channel-1/messages/msg-9/hostedContents/xyz/$value");
  });

  it("usa o path /messages/<pai>/replies/<id> para localizar resposta de thread e sua imagem embutida", async () => {
    const pngBytes = new Uint8Array([0x89, 0x50, 0x4e, 0x47]);
    const reply: GraphChatMessage = {
      id: "reply-2",
      createdDateTime: "2026-01-01T00:00:00Z",
      body: {
        contentType: "html",
        content:
          '<img src="https://graph.microsoft.com/v1.0/teams/team-1/channels/channel-1/messages/msg-9/replies/reply-2/hostedContents/abc/$value">',
      },
    };
    global.fetch = mock(async (input: RequestInfo | URL) => {
      const url = input.toString();
      requestedUrls.push(url);
      if (url.includes("/hostedContents/")) {
        return new Response(pngBytes, { status: 200, headers: { "content-type": "application/octet-stream" } });
      }
      return jsonResponse(reply);
    }) as unknown as typeof fetch;

    const teams = createTeamsService(fakeAuth, { downloadSharedFile: mock(async () => { throw new Error("nao deveria ser chamado"); }) });
    const location: TeamsMessageLocation = {
      kind: "channelReply", teamId: "team-1", channelId: "channel-1", parentMessageId: "msg-9", messageId: "reply-2",
    };

    const downloads = await teams.downloadMessageFiles(location);

    expect(downloads.length).toBe(1);
    expect(requestedUrls[0]).toBe("https://graph.microsoft.com/v1.0/teams/team-1/channels/channel-1/messages/msg-9/replies/reply-2");
    expect(requestedUrls[1]).toContain("/messages/msg-9/replies/reply-2/hostedContents/abc/$value");
  });

  it("lista respostas da thread pelo endpoint /replies da mensagem principal", async () => {
    global.fetch = mock(async (input: RequestInfo | URL) => {
      requestedUrls.push(input.toString());
      return jsonResponse({ value: [] });
    }) as unknown as typeof fetch;

    const teams = createTeamsService(fakeAuth, { downloadSharedFile: mock(async () => { throw new Error("nao deveria ser chamado"); }) });
    await teams.listChannelMessageReplies({ teamId: "team-1", channelId: "channel-1", messageId: "msg-9", top: 5 });

    expect(requestedUrls[0]).toBe("https://graph.microsoft.com/v1.0/teams/team-1/channels/channel-1/messages/msg-9/replies?$top=5");
  });
});
