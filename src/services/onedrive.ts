import type { AuthProvider } from "../types/auth.js";
import type {
  GraphDriveItem,
  GraphSharingLink,
  GraphPagedResponse,
} from "../types/graph.js";
import { graphFetch, graphFetchBinary } from "../utils/graph-client.js";
import {
  type ExtractResult,
  getFileExtension,
  isTextFile,
  isSupportedBinary,
  extractFromPdf,
  extractFromDocx,
} from "../utils/content-extractor.js";
import { SCOPES, DEFAULT_PAGE_SIZE_SMALL, DEFAULT_PAGE_SIZE_LARGE } from "../constants.js";
import { createGetToken } from "../utils/auth-helper.js";
import { flagName, isFlagGuardActive } from "../utils/flagged-paths.js";
import { assertDownloadable, storeDownload, type StoredDownload } from "../utils/download-store.js";

export interface ListFilesParams {
  itemId?: string;
  path?: string;
  top?: number;
}

export interface UploadFileParams {
  path: string;
  content: string;
}

export interface ReadSharedFileParams {
  driveId: string;
  itemId?: string;
  path?: string;
}

export interface ReadFileContentParams {
  fileName: string;
  startPage?: number;
  maxPages?: number;
}

export interface FileContentResult extends ExtractResult {
  displayName: string;
}

export interface SharedFileLocation {
  itemId?: string;
  path?: string;
}

export interface DriveFileLocation {
  driveId?: string;
  itemId?: string;
  path?: string;
}

interface DriveItemDownloadMetadata {
  name: string;
  size?: number;
  folder?: { childCount?: number };
  parentReference?: { path?: string };
}

export function buildDriveItemEndpoint(location: DriveFileLocation): string {
  const driveRoot = location.driveId ? `/drives/${location.driveId}` : "/me/drive";
  if (location.itemId) return `${driveRoot}/items/${location.itemId}`;
  if (location.path) return `${driveRoot}/root:/${location.path.replace(/^\//, "")}:`;
  throw new Error("É necessário informar itemId ou path do arquivo.");
}

export interface ShareFileParams {
  itemId: string;
  type?: "view" | "edit";
  scope?: "anonymous" | "organization";
}

interface ShareDriveItemResponse {
  id: string;
  name: string;
  parentReference?: { driveId?: string };
  webUrl?: string;
  file?: { mimeType?: string };
  folder?: { childCount?: number };
}

export function createOneDriveService(auth: AuthProvider) {
  const getToken = createGetToken(auth, SCOPES.DRIVE);

  async function listFiles(params: ListFilesParams = {}): Promise<GraphDriveItem[]> {
    const token = await getToken();
    const { itemId, path, top = DEFAULT_PAGE_SIZE_LARGE } = params;

    let endpoint: string;
    if (itemId) {
      endpoint = `/me/drive/items/${itemId}/children`;
    } else if (path) {
      endpoint = `/me/drive/root:/${path.replace(/^\//, "")}:/children`;
    } else {
      endpoint = "/me/drive/root/children";
    }

    const queryParams = new URLSearchParams({
      $top: String(top),
      $select: "id,name,size,lastModifiedDateTime,folder,file,webUrl,parentReference",
    });

    const result = await graphFetch<GraphPagedResponse<GraphDriveItem>>(
      token,
      `${endpoint}?${queryParams}`
    );
    return result.value.map((item) => ({
      ...item,
      name: flagName(item.name, item.parentReference?.path),
    }));
  }

  async function extractBinaryContent(
    contentUrl: string,
    params: ReadFileContentParams
  ): Promise<ExtractResult> {
    const token = await getToken();
    const ext = getFileExtension(params.fileName);
    const data = await graphFetchBinary(token, contentUrl);

    if (ext === ".pdf") {
      return extractFromPdf(data, params.startPage, params.maxPages);
    }
    if (ext === ".docx") {
      return extractFromDocx(data);
    }
    throw new Error(`Tipo de arquivo '${ext}' não suportado para extração de conteúdo. Suportados: .pdf, .docx e arquivos de texto.`);
  }

  async function resolveDisplayName(itemId: string, fileName: string): Promise<string> {
    if (!isFlagGuardActive()) return fileName;
    const token = await getToken();
    const item = await graphFetch<{ parentReference?: { path?: string } }>(
      token,
      `/me/drive/items/${itemId}?$select=parentReference`
    );
    return flagName(fileName, item.parentReference?.path);
  }

  async function readFileContent(
    itemId: string,
    params: ReadFileContentParams
  ): Promise<FileContentResult> {
    const displayName = await resolveDisplayName(itemId, params.fileName);
    const contentUrl = `/me/drive/items/${itemId}/content`;
    if (isTextFile(params.fileName)) {
      const token = await getToken();
      const text = await graphFetch<string>(token, contentUrl);
      return { text: typeof text === "string" ? text : JSON.stringify(text, null, 2), displayName };
    }
    if (isSupportedBinary(params.fileName)) {
      const result = await extractBinaryContent(contentUrl, params);
      return { ...result, displayName };
    }
    throw new Error(`Tipo de arquivo '${getFileExtension(params.fileName)}' não suportado. Suportados: .pdf, .docx e arquivos de texto.`);
  }

  async function resolveSharedDisplayName(
    driveId: string,
    fileName: string,
    location: SharedFileLocation
  ): Promise<string> {
    if (!isFlagGuardActive()) return fileName;
    if (location.path) return flagName(fileName, location.path);
    if (!location.itemId) return fileName;
    const token = await getToken();
    const item = await graphFetch<{ parentReference?: { path?: string } }>(
      token,
      `/drives/${driveId}/items/${location.itemId}?$select=parentReference`
    );
    return flagName(fileName, item.parentReference?.path);
  }

  async function readSharedFileContent(
    driveId: string,
    location: SharedFileLocation,
    params: ReadFileContentParams
  ): Promise<FileContentResult> {
    const contentUrl = buildSharedContentEndpoint(driveId, location.itemId, location.path);
    const displayName = await resolveSharedDisplayName(driveId, params.fileName, location);
    if (isTextFile(params.fileName)) {
      const token = await getToken();
      const text = await graphFetch<string>(token, contentUrl);
      return { text: typeof text === "string" ? text : JSON.stringify(text, null, 2), displayName };
    }
    if (isSupportedBinary(params.fileName)) {
      const result = await extractBinaryContent(contentUrl, params);
      return { ...result, displayName };
    }
    throw new Error(`Tipo de arquivo '${getFileExtension(params.fileName)}' não suportado. Suportados: .pdf, .docx e arquivos de texto.`);
  }

  function buildSharedContentEndpoint(driveId: string, itemId?: string, path?: string): string {
    return `${buildDriveItemEndpoint({ driveId, itemId, path })}/content`;
  }

  async function downloadDriveFile(location: DriveFileLocation): Promise<StoredDownload> {
    const token = await getToken();
    const itemEndpoint = buildDriveItemEndpoint(location);
    const item = await graphFetch<DriveItemDownloadMetadata>(
      token,
      `${itemEndpoint}?$select=name,size,folder,parentReference`
    );
    if (item.folder) throw new Error(`'${item.name}' é uma pasta; informe um arquivo.`);
    assertDownloadable(item.name, item.size ?? 0);
    const content = await graphFetchBinary(token, `${itemEndpoint}/content`);
    const stored = await storeDownload(item.name, content);
    return { ...stored, displayName: flagName(stored.fileName, item.parentReference?.path) };
  }

  async function resolveShareLink(shareUrl: string): Promise<{ driveId: string; itemId: string; name: string; webUrl?: string }> {
    const token = await getToken();
    const encoded = Buffer.from(shareUrl, "utf-8")
      .toString("base64")
      .replace(/\+/g, "-")
      .replace(/\//g, "_")
      .replace(/=+$/, "");
    const shareToken = `u!${encoded}`;

    const item = await graphFetch<ShareDriveItemResponse>(
      token,
      `/shares/${shareToken}/driveItem?$select=id,name,parentReference,webUrl,file,folder`
    );

    const driveId = item.parentReference?.driveId;
    if (!driveId) throw new Error("Não foi possível resolver o driveId do link compartilhado.");

    return { driveId, itemId: item.id, name: item.name, webUrl: item.webUrl };
  }

  async function uploadFile(params: UploadFileParams): Promise<GraphDriveItem> {
    const token = await getToken();
    const { path, content } = params;
    const cleanPath = path.replace(/^\//, "");

    return graphFetch<GraphDriveItem>(token, `/me/drive/root:/${cleanPath}:/content`, {
      method: "PUT",
      headers: { "Content-Type": "application/octet-stream" },
      body: content,
    });
  }

  async function searchFiles(
    query: string,
    top: number = DEFAULT_PAGE_SIZE_SMALL
  ): Promise<GraphDriveItem[]> {
    const token = await getToken();
    const queryParams = new URLSearchParams({
      $top: String(top),
      $select: "id,name,size,lastModifiedDateTime,webUrl,parentReference",
    });

    const result = await graphFetch<GraphPagedResponse<GraphDriveItem>>(
      token,
      `/me/drive/root/search(q='${encodeURIComponent(query)}')?${queryParams}`
    );
    return result.value.map((item) => ({
      ...item,
      name: flagName(item.name, item.parentReference?.path),
    }));
  }

  async function shareFile(params: ShareFileParams): Promise<GraphSharingLink> {
    const token = await getToken();
    const { itemId, type = "view", scope = "organization" } = params;

    return graphFetch<GraphSharingLink>(token, `/me/drive/items/${itemId}/createLink`, {
      method: "POST",
      body: JSON.stringify({ type, scope }),
    });
  }

  return {
    listFiles, readFileContent, readSharedFileContent, buildSharedContentEndpoint,
    downloadDriveFile, resolveShareLink, uploadFile, searchFiles, shareFile,
  };
}

export type OneDriveService = ReturnType<typeof createOneDriveService>;
