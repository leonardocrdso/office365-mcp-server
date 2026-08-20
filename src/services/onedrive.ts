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
      $select: "id,name,size,lastModifiedDateTime,folder,file,webUrl",
    });

    const result = await graphFetch<GraphPagedResponse<GraphDriveItem>>(
      token,
      `${endpoint}?${queryParams}`
    );
    return result.value;
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

  async function readFileContent(
    itemId: string,
    params: ReadFileContentParams
  ): Promise<ExtractResult> {
    const contentUrl = `/me/drive/items/${itemId}/content`;
    if (isTextFile(params.fileName)) {
      const token = await getToken();
      const text = await graphFetch<string>(token, contentUrl);
      return { text: typeof text === "string" ? text : JSON.stringify(text, null, 2) };
    }
    if (isSupportedBinary(params.fileName)) {
      return extractBinaryContent(contentUrl, params);
    }
    throw new Error(`Tipo de arquivo '${getFileExtension(params.fileName)}' não suportado. Suportados: .pdf, .docx e arquivos de texto.`);
  }

  async function readSharedFileContent(
    driveId: string,
    endpoint: string,
    params: ReadFileContentParams
  ): Promise<ExtractResult> {
    const contentUrl = endpoint;
    if (isTextFile(params.fileName)) {
      const token = await getToken();
      const text = await graphFetch<string>(token, contentUrl);
      return { text: typeof text === "string" ? text : JSON.stringify(text, null, 2) };
    }
    if (isSupportedBinary(params.fileName)) {
      return extractBinaryContent(contentUrl, params);
    }
    throw new Error(`Tipo de arquivo '${getFileExtension(params.fileName)}' não suportado. Suportados: .pdf, .docx e arquivos de texto.`);
  }

  function buildSharedContentEndpoint(driveId: string, itemId?: string, path?: string): string {
    if (itemId) return `/drives/${driveId}/items/${itemId}/content`;
    if (path) return `/drives/${driveId}/root:/${path.replace(/^\//, "")}:/content`;
    throw new Error("É necessário informar itemId ou path do arquivo.");
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
    return result.value;
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
    resolveShareLink, uploadFile, searchFiles, shareFile,
  };
}

export type OneDriveService = ReturnType<typeof createOneDriveService>;
