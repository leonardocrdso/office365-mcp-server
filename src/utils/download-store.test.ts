import { afterEach, beforeEach, describe, expect, it } from "bun:test";
import { mkdir, mkdtemp, readdir, readFile, rm, utimes, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { DOWNLOAD_MAX_BYTES, DOWNLOAD_RETENTION_MS } from "../constants.js";
import { assertDownloadable, sanitizeFileName, storeDownload, sweepExpiredDownloads } from "./download-store.js";

describe("sanitizeFileName", () => {
  it("remove tentativa de path traversal e mantem so o nome do arquivo", () => {
    expect(sanitizeFileName("../../etc/passwd")).toBe("passwd");
  });

  it("normaliza separador de caminho do windows para o nome do arquivo", () => {
    expect(sanitizeFileName("pasta\\sub\\caso.pdf")).toBe("caso.pdf");
  });

  it("substitui caracteres proibidos por underscore", () => {
    expect(sanitizeFileName("a:b?.pdf")).toBe("a_b_.pdf");
  });

  it("usa nome padrao quando o nome fica vazio", () => {
    expect(sanitizeFileName("")).toBe("arquivo");
  });

  it("usa nome padrao quando o nome e apenas '..'", () => {
    expect(sanitizeFileName("..")).toBe("arquivo");
  });

  it("mantem nomes ja validos inalterados", () => {
    expect(sanitizeFileName("05 Module Overview (Alcacer).pdf")).toBe("05 Module Overview (Alcacer).pdf");
  });
});

describe("assertDownloadable", () => {
  it("lanca erro quando o tamanho excede o limite", () => {
    expect(() => assertDownloadable("arquivo.pdf", DOWNLOAD_MAX_BYTES + 1)).toThrow();
  });

  it("nao lanca erro quando o tamanho e exatamente o limite", () => {
    expect(() => assertDownloadable("arquivo.pdf", DOWNLOAD_MAX_BYTES)).not.toThrow();
  });
});

describe("storeDownload", () => {
  let originalHome: string | undefined;
  let tempHome: string;

  beforeEach(async () => {
    originalHome = process.env.OFFICE365_MCP_HOME;
    tempHome = await mkdtemp(join(tmpdir(), "o365-test-"));
    process.env.OFFICE365_MCP_HOME = tempHome;
  });

  afterEach(async () => {
    if (originalHome === undefined) delete process.env.OFFICE365_MCP_HOME;
    else process.env.OFFICE365_MCP_HOME = originalHome;
    await rm(tempHome, { recursive: true, force: true });
  });

  it("grava o conteudo em downloads/<uuid>/<nome> e retorna metadados coerentes", async () => {
    const content = new TextEncoder().encode("conteudo de teste");

    const stored = await storeDownload("caso.pdf", content);

    const downloadsRoot = join(tempHome, "downloads");
    const [uuidDir] = await readdir(downloadsRoot);
    expect(uuidDir).toBeDefined();
    if (!uuidDir) return;
    expect(stored.path).toBe(join(downloadsRoot, uuidDir, "caso.pdf"));
    expect(stored.fileName).toBe("caso.pdf");
    expect(stored.displayName).toBe("caso.pdf");
    expect(stored.sizeBytes).toBe(content.byteLength);

    const written = await readFile(stored.path);
    expect(new Uint8Array(written)).toEqual(content);
  });

  it("mantem apenas o basename do nome original, descartando qualquer subcaminho", async () => {
    const content = new TextEncoder().encode("outro conteudo");

    const stored = await storeDownload("materiais/turma/caso.pdf", content);

    expect(stored.fileName).toBe("caso.pdf");
    const written = await readFile(stored.path);
    expect(new Uint8Array(written)).toEqual(content);
  });

  it("remove subdiretorio expirado e mantem um recente na proxima chamada", async () => {
    const downloadsRoot = join(tempHome, "downloads");
    await mkdir(downloadsRoot, { recursive: true });

    const oldDir = join(downloadsRoot, "old-download");
    await mkdir(oldDir, { recursive: true });
    await writeFile(join(oldDir, "velho.pdf"), "conteudo antigo");
    const expiredTime = new Date(Date.now() - DOWNLOAD_RETENTION_MS - 1_000);
    await utimes(oldDir, expiredTime, expiredTime);

    const recentDir = join(downloadsRoot, "recent-download");
    await mkdir(recentDir, { recursive: true });
    await writeFile(join(recentDir, "novo.pdf"), "conteudo novo");

    await storeDownload("outro.pdf", new TextEncoder().encode("mais um"));

    const remaining = await readdir(downloadsRoot);
    expect(remaining).toContain("recent-download");
    expect(remaining).not.toContain("old-download");
  });

  it("varredura sem novo download apaga o expirado e mantem o recente", async () => {
    const downloadsRoot = join(tempHome, "downloads");
    const oldDir = join(downloadsRoot, "ja-entregue");
    const recentDir = join(downloadsRoot, "acabou-de-baixar");
    await mkdir(oldDir, { recursive: true });
    await mkdir(recentDir, { recursive: true });
    const expiredTime = new Date(Date.now() - DOWNLOAD_RETENTION_MS - 1_000);
    await utimes(oldDir, expiredTime, expiredTime);

    await sweepExpiredDownloads();

    expect(await readdir(downloadsRoot)).toEqual(["acabou-de-baixar"]);
  });

  it("varredura sem pasta de downloads nao falha", async () => {
    await expect(sweepExpiredDownloads()).resolves.toBeUndefined();
  });
});
