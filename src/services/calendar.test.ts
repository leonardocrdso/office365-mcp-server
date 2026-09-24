import { afterEach, beforeEach, describe, expect, it, mock } from "bun:test";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import type { AuthProvider } from "../types/auth.js";
import type { GraphMessageAttachment } from "../types/graph.js";
import { createCalendarService } from "./calendar.js";

const fakeAuth: AuthProvider = {
  async getAccessToken() {
    return "fake-token";
  },
};

describe("downloadEventAttachments", () => {
  const originalFetch = global.fetch;
  let originalHome: string | undefined;
  let tempHome: string;
  let requestedUrls: string[];

  const attachmentId = "AAMkAGI1AAA=/AAA=";
  const conteudoBinario = new Uint8Array([5, 6, 7, 8]);

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

  it("baixa o anexo do evento, grava em disco e monta a url com o id codificado", async () => {
    const attachment: GraphMessageAttachment = {
      id: attachmentId,
      name: "pauta.pdf",
      contentType: "application/pdf",
      size: conteudoBinario.byteLength,
      isInline: false,
      "@odata.type": "#microsoft.graph.fileAttachment",
    };
    mockListingThenBinary([attachment]);

    const calendar = createCalendarService(fakeAuth);
    const downloads = await calendar.downloadEventAttachments("evt-1");

    expect(downloads.length).toBe(1);
    const [download] = downloads;
    expect(download).toBeDefined();
    if (!download) return;
    expect(download.fileName).toBe("pauta.pdf");

    const written = await readFile(download.path);
    expect(new Uint8Array(written)).toEqual(conteudoBinario);

    expect(requestedUrls[0]).toContain("/me/events/evt-1/attachments?");
    expect(requestedUrls[1]).toContain(
      `/me/events/evt-1/attachments/${encodeURIComponent(attachmentId)}/$value`
    );
  });
});
