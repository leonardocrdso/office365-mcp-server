import { msalClient } from "../auth/msal-client.ts";
import { graphFetch } from "../utils/errors.ts";

const DRIVE_SCOPES = ["Files.Read.All", "Files.ReadWrite.All"];

async function getToken() {
  return msalClient.getAccessToken(DRIVE_SCOPES);
}

export interface ListFilesParams {
  itemId?: string;
  path?: string;
  top?: number;
}

export async function listFiles(params: ListFilesParams = {}) {
  const token = await getToken();
  const { itemId, path, top = 20 } = params;

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

  const result = await graphFetch(token, `${endpoint}?${queryParams}`);
  return result.value;
}

export async function readFileContent(itemId: string) {
  const token = await getToken();
  const content = await graphFetch(token, `/me/drive/items/${itemId}/content`);
  return content;
}

export interface UploadFileParams {
  path: string;
  content: string;
}

export async function uploadFile(params: UploadFileParams) {
  const token = await getToken();
  const { path, content } = params;
  const cleanPath = path.replace(/^\//, "");

  const result = await graphFetch(token, `/me/drive/root:/${cleanPath}:/content`, {
    method: "PUT",
    headers: { "Content-Type": "application/octet-stream" },
    body: content,
  });
  return result;
}

export async function searchFiles(query: string, top: number = 10) {
  const token = await getToken();
  const queryParams = new URLSearchParams({
    $top: String(top),
    $select: "id,name,size,lastModifiedDateTime,webUrl,parentReference",
  });

  const result = await graphFetch(
    token,
    `/me/drive/root/search(q='${encodeURIComponent(query)}')?${queryParams}`
  );
  return result.value;
}

export interface ShareFileParams {
  itemId: string;
  type?: "view" | "edit";
  scope?: "anonymous" | "organization";
}

export async function shareFile(params: ShareFileParams) {
  const token = await getToken();
  const { itemId, type = "view", scope = "organization" } = params;

  const result = await graphFetch(token, `/me/drive/items/${itemId}/createLink`, {
    method: "POST",
    body: JSON.stringify({ type, scope }),
  });
  return result;
}
