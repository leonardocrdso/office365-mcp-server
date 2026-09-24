import { afterEach, beforeEach, describe, expect, it, mock } from "bun:test";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import type { AuthProvider } from "../types/auth.js";
import { buildDriveItemEndpoint, createOneDriveService } from "./onedrive.js";

const FLAG_ENV_VAR = "OFFICE365_MCP_FLAG_PATH_MARKERS";
const FLAG_LABEL = "[MATERIAL DE OUTRA TURMA/PASTA — confira data e conteúdo antes de citar] ";

const fakeAuth: AuthProvider = {
  async getAccessToken() {
    return "fake-token";
  },
};

const nomeArquivoTurmaAnterior = "05 Module Overview - Strategy (Alcacer) OPM64.3.pdf";
const caminhoTurmaAnterior = "/drive/root:/Documentos/Gabriel Sens/Harvard/OPM 64.3";
const nomeArquivoTurmaAtual = "05 Module Overview - Strategy (Alcacer) OPM65.3.pdf";
const caminhoTurmaAtual = "/drive/root:/Documentos/Gabriel Sens/Harvard/OPM 65.3";

function jsonResponse(body: unknown): Response {
  return new Response(JSON.stringify(body), {
    status: 200,
    headers: { "content-type": "application/json" },
  });
}

describe("onedrive - marcacao de material de outra turma/pasta", () => {
  const originalFetch = global.fetch;
  const originalEnv = process.env[FLAG_ENV_VAR];
  let requestedUrls: string[];

  beforeEach(() => {
    requestedUrls = [];
  });

  afterEach(() => {
    global.fetch = originalFetch;
    if (originalEnv === undefined) delete process.env[FLAG_ENV_VAR];
    else process.env[FLAG_ENV_VAR] = originalEnv;
  });

  function mockDriveListing(items: unknown[]): void {
    global.fetch = mock(async (input: RequestInfo | URL) => {
      requestedUrls.push(input.toString());
      return jsonResponse({ value: items });
    }) as unknown as typeof fetch;
  }

  it("marca o nome do arquivo quando search-files encontra material da turma configurada", async () => {
    process.env[FLAG_ENV_VAR] = "OPM 64.3,OPM64.3";
    mockDriveListing([
      {
        id: "1",
        name: nomeArquivoTurmaAnterior,
        size: 100,
        webUrl: "https://contoso.sharepoint.com/x",
        parentReference: { path: caminhoTurmaAnterior },
      },
      {
        id: "2",
        name: nomeArquivoTurmaAtual,
        size: 100,
        webUrl: "https://contoso.sharepoint.com/y",
        parentReference: { path: caminhoTurmaAtual },
      },
    ]);

    const onedrive = createOneDriveService(fakeAuth);
    const items = await onedrive.searchFiles("Module Overview Strategy");

    expect(items[0]!.name).toBe(`${FLAG_LABEL}${nomeArquivoTurmaAnterior}`);
    expect(items[1]!.name).toBe(nomeArquivoTurmaAtual);
  });

  it("marca o nome do arquivo quando list-drive-files lista a pasta da turma configurada", async () => {
    process.env[FLAG_ENV_VAR] = "OPM 64.3,OPM64.3";
    mockDriveListing([
      {
        id: "1",
        name: nomeArquivoTurmaAnterior,
        lastModifiedDateTime: "2026-01-01T00:00:00Z",
        webUrl: "https://contoso.sharepoint.com/x",
        parentReference: { path: caminhoTurmaAnterior },
      },
    ]);

    const onedrive = createOneDriveService(fakeAuth);
    const items = await onedrive.listFiles({ path: "Documentos/Gabriel Sens/Harvard/OPM 64.3" });

    expect(items[0]!.name).toBe(`${FLAG_LABEL}${nomeArquivoTurmaAnterior}`);
  });

  it("nao marca nada quando a variavel de ambiente nao esta configurada", async () => {
    delete process.env[FLAG_ENV_VAR];
    mockDriveListing([
      {
        id: "1",
        name: nomeArquivoTurmaAnterior,
        size: 100,
        webUrl: "https://contoso.sharepoint.com/x",
        parentReference: { path: caminhoTurmaAnterior },
      },
    ]);

    const onedrive = createOneDriveService(fakeAuth);
    const items = await onedrive.searchFiles("Module Overview Strategy");

    expect(items[0]!.name).toBe(nomeArquivoTurmaAnterior);
  });

  it("read-file-content marca o displayName sem alterar o texto extraido, quando a guarda esta ativa", async () => {
    process.env[FLAG_ENV_VAR] = "OPM 64.3,OPM64.3";
    const conteudo = "Conteudo do caso KLog.co";
    global.fetch = mock(async (input: RequestInfo | URL) => {
      const url = input.toString();
      requestedUrls.push(url);
      if (url.includes("$select=parentReference")) {
        return jsonResponse({ parentReference: { path: caminhoTurmaAnterior } });
      }
      return new Response(conteudo, { status: 200, headers: { "content-type": "text/plain" } });
    }) as unknown as typeof fetch;

    const onedrive = createOneDriveService(fakeAuth);
    const result = await onedrive.readFileContent("item-1", {
      fileName: "05 Module Overview - Strategy (Alcacer) OPM64.3.md",
    });

    expect(result.text).toBe(conteudo);
    expect(result.displayName).toBe(`${FLAG_LABEL}05 Module Overview - Strategy (Alcacer) OPM64.3.md`);
    expect(requestedUrls.some((url) => url.includes("$select=parentReference"))).toBe(true);
  });

  it("read-file-content nao faz chamada extra quando a guarda nao esta configurada", async () => {
    delete process.env[FLAG_ENV_VAR];
    const conteudo = "Conteudo qualquer";
    global.fetch = mock(async (input: RequestInfo | URL) => {
      requestedUrls.push(input.toString());
      return new Response(conteudo, { status: 200, headers: { "content-type": "text/plain" } });
    }) as unknown as typeof fetch;

    const onedrive = createOneDriveService(fakeAuth);
    const result = await onedrive.readFileContent("item-1", { fileName: "notas.md" });

    expect(result.text).toBe(conteudo);
    expect(result.displayName).toBe("notas.md");
    expect(requestedUrls.length).toBe(1);
  });

  it("read-shared-file-content marca o displayName buscando o metadado quando so o itemId e informado, como no uso real com driveId+itemId", async () => {
    process.env[FLAG_ENV_VAR] = "OPM 64.3,OPM64.3";
    const conteudo = "ResponseSummary do modulo";
    global.fetch = mock(async (input: RequestInfo | URL) => {
      const url = input.toString();
      requestedUrls.push(url);
      if (url.includes("$select=parentReference")) {
        return jsonResponse({ parentReference: { path: caminhoTurmaAnterior } });
      }
      return new Response(conteudo, { status: 200, headers: { "content-type": "text/plain" } });
    }) as unknown as typeof fetch;

    const onedrive = createOneDriveService(fakeAuth);
    const result = await onedrive.readSharedFileContent(
      "drive-1",
      { itemId: "item-shared-1" },
      { fileName: "ResponseSummary. OPM64.3 Unit 3.md" }
    );

    expect(result.text).toBe(conteudo);
    expect(result.displayName).toBe(`${FLAG_LABEL}ResponseSummary. OPM64.3 Unit 3.md`);
    expect(requestedUrls.some((url) => url.includes("/drives/drive-1/items/item-shared-1?$select=parentReference"))).toBe(true);
  });

  it("read-shared-file-content marca o displayName pelo path informado, sem chamada de metadado extra", async () => {
    process.env[FLAG_ENV_VAR] = "OPM 64.3,OPM64.3";
    const conteudo = "Conteudo do caso";
    global.fetch = mock(async (input: RequestInfo | URL) => {
      requestedUrls.push(input.toString());
      return new Response(conteudo, { status: 200, headers: { "content-type": "text/plain" } });
    }) as unknown as typeof fetch;

    const onedrive = createOneDriveService(fakeAuth);
    const result = await onedrive.readSharedFileContent(
      "drive-1",
      { path: "Documentos/Gabriel Sens/Harvard/OPM 64.3/notas.md" },
      { fileName: "notas.md" }
    );

    expect(result.text).toBe(conteudo);
    expect(result.displayName).toBe(`${FLAG_LABEL}notas.md`);
    expect(requestedUrls.length).toBe(1);
  });

  it("read-shared-file-content nao faz chamada extra quando a guarda nao esta configurada", async () => {
    delete process.env[FLAG_ENV_VAR];
    const conteudo = "Conteudo qualquer";
    global.fetch = mock(async (input: RequestInfo | URL) => {
      requestedUrls.push(input.toString());
      return new Response(conteudo, { status: 200, headers: { "content-type": "text/plain" } });
    }) as unknown as typeof fetch;

    const onedrive = createOneDriveService(fakeAuth);
    const result = await onedrive.readSharedFileContent(
      "drive-1",
      { itemId: "item-shared-1" },
      { fileName: "ResponseSummary. OPM65.3 Unit 3.md" }
    );

    expect(result.text).toBe(conteudo);
    expect(result.displayName).toBe("ResponseSummary. OPM65.3 Unit 3.md");
    expect(requestedUrls.length).toBe(1);
  });
});

describe("buildDriveItemEndpoint", () => {
  it("monta endpoint pelo drive pessoal quando so o itemId e informado", () => {
    expect(buildDriveItemEndpoint({ itemId: "X" })).toBe("/me/drive/items/X");
  });

  it("monta endpoint por driveId e itemId quando ambos sao informados", () => {
    expect(buildDriveItemEndpoint({ driveId: "D", itemId: "X" })).toBe("/drives/D/items/X");
  });

  it("monta endpoint por path com driveId, envolvendo o caminho entre dois-pontos", () => {
    expect(buildDriveItemEndpoint({ driveId: "D", path: "/Docs/a b.pdf" })).toBe(
      "/drives/D/root:/Docs/a b.pdf:"
    );
  });

  it("lanca erro quando nem itemId nem path sao informados", () => {
    expect(() => buildDriveItemEndpoint({})).toThrow();
  });
});

describe("downloadDriveFile", () => {
  const originalFetch = global.fetch;
  let originalHome: string | undefined;
  let originalFlag: string | undefined;
  let tempHome: string;
  let requestedUrls: string[];

  beforeEach(async () => {
    originalHome = process.env.OFFICE365_MCP_HOME;
    originalFlag = process.env[FLAG_ENV_VAR];
    tempHome = await mkdtemp(join(tmpdir(), "o365-test-"));
    process.env.OFFICE365_MCP_HOME = tempHome;
    requestedUrls = [];
  });

  afterEach(async () => {
    global.fetch = originalFetch;
    if (originalHome === undefined) delete process.env.OFFICE365_MCP_HOME;
    else process.env.OFFICE365_MCP_HOME = originalHome;
    if (originalFlag === undefined) delete process.env[FLAG_ENV_VAR];
    else process.env[FLAG_ENV_VAR] = originalFlag;
    await rm(tempHome, { recursive: true, force: true });
  });

  function mockMetadataThenContent(metadata: unknown, content: Uint8Array<ArrayBuffer>): void {
    global.fetch = mock(async (input: RequestInfo | URL) => {
      const url = input.toString();
      requestedUrls.push(url);
      if (url.includes("$select=name,size,folder,parentReference")) {
        return jsonResponse(metadata);
      }
      return new Response(content, { status: 200, headers: { "content-type": "application/octet-stream" } });
    }) as unknown as typeof fetch;
  }

  it("baixa o metadado, depois o conteudo, e grava o arquivo em disco", async () => {
    const conteudo = new Uint8Array([9, 8, 7]);
    mockMetadataThenContent({ name: "caso.pdf", size: conteudo.byteLength }, conteudo);

    const onedrive = createOneDriveService(fakeAuth);
    const stored = await onedrive.downloadDriveFile({ itemId: "item-1" });

    expect(stored.fileName).toBe("caso.pdf");
    const written = await readFile(stored.path);
    expect(new Uint8Array(written)).toEqual(conteudo);
    expect(requestedUrls[0]).toContain("$select=name,size,folder,parentReference");
    expect(requestedUrls[1]).toContain("/me/drive/items/item-1/content");
  });

  it("lanca erro quando o item e uma pasta, sem chamar /content", async () => {
    global.fetch = mock(async (input: RequestInfo | URL) => {
      requestedUrls.push(input.toString());
      return jsonResponse({ name: "Materiais", folder: { childCount: 3 } });
    }) as unknown as typeof fetch;

    const onedrive = createOneDriveService(fakeAuth);

    await expect(onedrive.downloadDriveFile({ itemId: "item-pasta" })).rejects.toThrow();
    expect(requestedUrls.length).toBe(1);
  });

  it("com flag de turma antiga ativa, marca o displayName mas mantem fileName e o nome em disco limpos", async () => {
    process.env[FLAG_ENV_VAR] = "OPM 64";
    const conteudo = new Uint8Array([1]);
    mockMetadataThenContent(
      {
        name: "caso.pdf",
        size: conteudo.byteLength,
        parentReference: { path: "/drive/root:/Documentos/Harvard/OPM 64.3" },
      },
      conteudo
    );

    const onedrive = createOneDriveService(fakeAuth);
    const stored = await onedrive.downloadDriveFile({ itemId: "item-2" });

    expect(stored.fileName).toBe("caso.pdf");
    expect(stored.displayName).toBe(`${FLAG_LABEL}caso.pdf`);
    expect(stored.path.endsWith("/caso.pdf")).toBe(true);
  });
});

describe("downloadSharedFile", () => {
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

  it("resolve o link via /shares e baixa o conteudo do item resolvido", async () => {
    const shareUrl = "https://contoso.sharepoint.com/:w:/s/abc/xyz";
    const conteudo = new Uint8Array([9, 9, 9]);

    global.fetch = mock(async (input: RequestInfo | URL) => {
      const url = input.toString();
      requestedUrls.push(url);
      if (url.includes("/shares/")) {
        return jsonResponse({ id: "item-1", name: "contrato.pdf", parentReference: { driveId: "drive-1" } });
      }
      if (url.includes("$select=name,size,folder,parentReference")) {
        return jsonResponse({ name: "contrato.pdf", size: conteudo.byteLength });
      }
      return new Response(conteudo, { status: 200, headers: { "content-type": "application/octet-stream" } });
    }) as unknown as typeof fetch;

    const onedrive = createOneDriveService(fakeAuth);
    const stored = await onedrive.downloadSharedFile(shareUrl);

    const expectedToken = Buffer.from(shareUrl, "utf-8")
      .toString("base64")
      .replace(/\+/g, "-")
      .replace(/\//g, "_")
      .replace(/=+$/, "");

    expect(requestedUrls[0]).toContain(`/shares/u!${expectedToken}/driveItem`);
    expect(requestedUrls[1]).toContain("/drives/drive-1/items/item-1");
    expect(stored.fileName).toBe("contrato.pdf");

    const written = await readFile(stored.path);
    expect(new Uint8Array(written)).toEqual(conteudo);
  });
});
