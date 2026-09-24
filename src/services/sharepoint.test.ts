import { afterEach, describe, expect, it, mock } from "bun:test";
import type { AuthProvider } from "../types/auth.js";
import { createSharePointService } from "./sharepoint.js";

const FLAG_ENV_VAR = "OFFICE365_MCP_FLAG_PATH_MARKERS";
const FLAG_LABEL = "[MATERIAL DE OUTRA TURMA/PASTA — confira data e conteúdo antes de citar] ";

const fakeAuth: AuthProvider = {
  async getAccessToken() {
    return "fake-token";
  },
};

const nomeArquivoTurmaAnterior = "05 Module Overview - Strategy (Alcacer) OPM64.3.pdf";
const urlTurmaAnterior =
  "https://contoso.sharepoint.com/sites/Gabriel/Documentos/Gabriel%20Sens/Harvard/OPM%2064.3/05%20Module%20Overview%20-%20Strategy%20(Alcacer)%20OPM64.3.pdf";
const nomeArquivoTurmaAtual = "05 Module Overview - Strategy (Alcacer) OPM65.3.pdf";
const urlTurmaAtual =
  "https://contoso.sharepoint.com/sites/Gabriel/Documentos/Gabriel%20Sens/Harvard/OPM%2065.3/05%20Module%20Overview%20-%20Strategy%20(Alcacer)%20OPM65.3.pdf";
const caminhoTurmaAnterior = "/drives/drive-1/root:/Documentos/Gabriel Sens/Harvard/OPM 64.3";
const caminhoTurmaAtual = "/drives/drive-1/root:/Documentos/Gabriel Sens/Harvard/OPM 65.3";

function jsonResponse(body: unknown): Response {
  return new Response(JSON.stringify(body), {
    status: 200,
    headers: { "content-type": "application/json" },
  });
}

function mockSearchResponse(): void {
  global.fetch = mock(async () =>
    new Response(
      JSON.stringify({
        value: [
          {
            hitsContainers: [
              {
                hits: [
                  {
                    hitId: "1",
                    resource: {
                      "@odata.type": "#microsoft.graph.driveItem",
                      name: nomeArquivoTurmaAnterior,
                      webUrl: urlTurmaAnterior,
                      parentReference: { driveId: "drive-1" },
                    },
                  },
                  {
                    hitId: "2",
                    resource: {
                      "@odata.type": "#microsoft.graph.driveItem",
                      name: nomeArquivoTurmaAtual,
                      webUrl: urlTurmaAtual,
                      parentReference: { driveId: "drive-1" },
                    },
                  },
                ],
              },
            ],
          },
        ],
      }),
      { status: 200, headers: { "content-type": "application/json" } }
    )
  ) as unknown as typeof fetch;
}

describe("sharepoint - marcacao de material de outra turma/pasta", () => {
  const originalFetch = global.fetch;
  const originalEnv = process.env[FLAG_ENV_VAR];

  afterEach(() => {
    global.fetch = originalFetch;
    if (originalEnv === undefined) delete process.env[FLAG_ENV_VAR];
    else process.env[FLAG_ENV_VAR] = originalEnv;
  });

  it("marca o nome do driveItem quando search-sharepoint encontra material da turma configurada", async () => {
    process.env[FLAG_ENV_VAR] = "OPM 64.3,OPM64.3";
    mockSearchResponse();

    const sharepoint = createSharePointService(fakeAuth);
    const [response] = await sharepoint.searchSharePoint("Module Overview Strategy");
    const hits = response!.hitsContainers[0]!.hits;

    expect(hits[0]!.resource.name).toBe(`${FLAG_LABEL}${nomeArquivoTurmaAnterior}`);
    expect(hits[1]!.resource.name).toBe(nomeArquivoTurmaAtual);
  });

  it("nao marca nada quando a variavel de ambiente nao esta configurada", async () => {
    delete process.env[FLAG_ENV_VAR];
    mockSearchResponse();

    const sharepoint = createSharePointService(fakeAuth);
    const [response] = await sharepoint.searchSharePoint("Module Overview Strategy");
    const hits = response!.hitsContainers[0]!.hits;

    expect(hits[0]!.resource.name).toBe(nomeArquivoTurmaAnterior);
    expect(hits[1]!.resource.name).toBe(nomeArquivoTurmaAtual);
  });

  it("marca o nome do item quando list-library-items lista a pasta da turma configurada", async () => {
    process.env[FLAG_ENV_VAR] = "OPM 64.3,OPM64.3";
    global.fetch = mock(async () =>
      jsonResponse({
        value: [
          {
            id: "1",
            name: nomeArquivoTurmaAnterior,
            webUrl: urlTurmaAnterior,
            parentReference: { path: caminhoTurmaAnterior },
          },
          {
            id: "2",
            name: nomeArquivoTurmaAtual,
            webUrl: urlTurmaAtual,
            parentReference: { path: caminhoTurmaAtual },
          },
        ],
      })
    ) as unknown as typeof fetch;

    const sharepoint = createSharePointService(fakeAuth);
    const items = await sharepoint.listLibraryItems({ driveId: "drive-1" });

    expect(items[0]!.name).toBe(`${FLAG_LABEL}${nomeArquivoTurmaAnterior}`);
    expect(items[1]!.name).toBe(nomeArquivoTurmaAtual);
  });

  it("list-library-items nao marca nada quando a variavel de ambiente nao esta configurada", async () => {
    delete process.env[FLAG_ENV_VAR];
    global.fetch = mock(async () =>
      jsonResponse({
        value: [
          {
            id: "1",
            name: nomeArquivoTurmaAnterior,
            webUrl: urlTurmaAnterior,
            parentReference: { path: caminhoTurmaAnterior },
          },
        ],
      })
    ) as unknown as typeof fetch;

    const sharepoint = createSharePointService(fakeAuth);
    const items = await sharepoint.listLibraryItems({ driveId: "drive-1" });

    expect(items[0]!.name).toBe(nomeArquivoTurmaAnterior);
  });
});
