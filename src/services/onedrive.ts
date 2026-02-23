import type { AuthProvider } from "../types/auth.js";
import type {
  GraphDriveItem,
  GraphSharingLink,
  GraphPagedResponse,
} from "../types/graph.js";
import { graphFetch } from "../utils/errors.js";
import { SCOPES, DEFAULT_PAGE_SIZE_SMALL, DEFAULT_PAGE_SIZE_LARGE } from "../constants.js";

export interface ListFilesParams {
  itemId?: string;
  path?: string;
  top?: number;
}

export interface UploadFileParams {
  path: string;
  content: string;
}

export interface ShareFileParams {
  itemId: string;
  type?: "view" | "edit";
  scope?: "anonymous" | "organization";
}

export function createOneDriveService(auth: AuthProvider) {
  async function getToken() {
    return auth.getAccessToken([...SCOPES.DRIVE]);
  }

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

  async function readFileContent(itemId: string): Promise<string> {
    const token = await getToken();
    return graphFetch<string>(token, `/me/drive/items/${itemId}/content`);
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

  return { listFiles, readFileContent, uploadFile, searchFiles, shareFile };
}

export type OneDriveService = ReturnType<typeof createOneDriveService>;
