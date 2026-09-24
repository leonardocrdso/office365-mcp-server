import { afterEach, beforeEach, describe, expect, it, mock } from "bun:test";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { DOWNLOAD_MAX_BYTES } from "../constants.js";
import type { AuthProvider } from "../types/auth.js";
import type { GraphMessageAttachment } from "../types/graph.js";
import { createMailService } from "./mail.js";

const fakeAuth: AuthProvider = {
  async getAccessToken() {
    return "fake-token";
  },
};

function emptyMessagesPage(): Response {
  return new Response(JSON.stringify({ value: [] }), {
    status: 200,
    headers: { "content-type": "application/json" },
  });
}

function decodedSearchParam(requestedUrl: string): string {
  const url = new URL(requestedUrl, "https://graph.microsoft.com");
  return url.searchParams.get("$search") ?? "";
}

describe("searchEmails - montagem do $search KQL", () => {
  let requestedUrl: string | undefined;
  const originalFetch = global.fetch;

  beforeEach(() => {
    requestedUrl = undefined;
    global.fetch = mock(async (input: RequestInfo | URL) => {
      requestedUrl = input.toString();
      return emptyMessagesPage();
    }) as unknown as typeof fetch;
  });

  afterEach(() => {
    global.fetch = originalFetch;
  });

  const consultasReaisQueQuebravamOMcp: Array<{ query: string; fraseEsperada: string }> = [
    { query: '"October 15"', fraseEsperada: '"\\"October 15\\""' },
    { query: '"Axel Springer"', fraseEsperada: '"\\"Axel Springer\\""' },
    { query: '"Individual Strategy"', fraseEsperada: '"\\"Individual Strategy\\""' },
    { query: '"OPM 65"', fraseEsperada: '"\\"OPM 65\\""' },
    { query: "11/10", fraseEsperada: '"\\"11/10\\""' },
  ];

  for (const { query, fraseEsperada } of consultasReaisQueQuebravamOMcp) {
    it(`monta $search valido para a consulta real ${JSON.stringify(query)}`, async () => {
      const mail = createMailService(fakeAuth);
      await mail.searchEmails(query);

      const search = decodedSearchParam(requestedUrl!);
      expect(search).toBe(fraseEsperada);
      expect(search.startsWith('""')).toBe(false);
    });
  }

  it("escapa aspas internas sem duplicar a delimitacao externa", async () => {
    const mail = createMailService(fakeAuth);
    await mail.searchEmails('diga "oi" para o cliente');

    const search = decodedSearchParam(requestedUrl!);
    expect(search).toBe('"\\"diga \\"oi\\" para o cliente\\""');
    expect(search.startsWith('""')).toBe(false);
  });
});

describe("downloadAttachments", () => {
  const originalFetch = global.fetch;
  let originalHome: string | undefined;
  let tempHome: string;
  let requestedUrls: string[];

  const attachmentId = "AAMkAGI1AAA=/AAA=";
  const conteudoBinario = new Uint8Array([1, 2, 3, 4]);

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

  function mockListingThenBinary(listing: readonly GraphMessageAttachment[]): void {
    global.fetch = mock(async (input: RequestInfo | URL) => {
      const url = input.toString();
      requestedUrls.push(url);
      if (url.endsWith("/$value")) {
        return new Response(conteudoBinario, {
          status: 200,
          headers: { "content-type": "application/octet-stream" },
        });
      }
      return new Response(JSON.stringify({ value: listing }), {
        status: 200,
        headers: { "content-type": "application/json" },
      });
    }) as unknown as typeof fetch;
  }

  it("baixa o anexo selecionado, grava em disco e monta a url com o id codificado", async () => {
    const attachment: GraphMessageAttachment = {
      id: attachmentId,
      name: "caso.pdf",
      contentType: "application/pdf",
      size: conteudoBinario.byteLength,
      isInline: false,
      "@odata.type": "#microsoft.graph.fileAttachment",
    };
    mockListingThenBinary([attachment]);

    const mail = createMailService(fakeAuth);
    const downloads = await mail.downloadAttachments("msg-1");

    expect(downloads.length).toBe(1);
    const [download] = downloads;
    expect(download).toBeDefined();
    if (!download) return;
    expect(download.fileName).toBe("caso.pdf");

    const written = await readFile(download.path);
    expect(new Uint8Array(written)).toEqual(conteudoBinario);

    expect(requestedUrls[0]).toContain("/me/messages/msg-1/attachments?");
    expect(requestedUrls[1]).toContain(
      `/me/messages/msg-1/attachments/${encodeURIComponent(attachmentId)}/$value`
    );
  });

  it("nao baixa nada quando um anexo selecionado excede o limite de tamanho informado na listagem", async () => {
    const attachmentGrande: GraphMessageAttachment = {
      id: "id-grande",
      name: "video.mp4",
      contentType: "video/mp4",
      size: DOWNLOAD_MAX_BYTES + 1,
      isInline: false,
      "@odata.type": "#microsoft.graph.fileAttachment",
    };
    mockListingThenBinary([attachmentGrande]);

    const mail = createMailService(fakeAuth);

    await expect(mail.downloadAttachments("msg-1")).rejects.toThrow();
    expect(requestedUrls.length).toBe(1);
  });
});
