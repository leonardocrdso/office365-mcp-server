import { msalClient } from "../auth/msal-client.ts";
import { graphFetch } from "../utils/errors.ts";

const SP_SCOPES = ["Sites.Read.All", "Sites.ReadWrite.All"];

async function getToken() {
  return msalClient.getAccessToken(SP_SCOPES);
}

export async function listSites(query?: string) {
  const token = await getToken();
  const searchQuery = query ?? "*";
  const result = await graphFetch(
    token,
    `/sites?search=${encodeURIComponent(searchQuery)}&$select=id,displayName,name,webUrl,description`
  );
  return result.value;
}

export async function getSite(siteId: string) {
  const token = await getToken();
  const result = await graphFetch(
    token,
    `/sites/${siteId}?$select=id,displayName,name,webUrl,description`
  );
  return result;
}

export async function listDocumentLibraries(siteId: string) {
  const token = await getToken();
  const result = await graphFetch(
    token,
    `/sites/${siteId}/drives?$select=id,name,driveType,webUrl,description`
  );
  return result.value;
}

export interface ListLibraryItemsParams {
  driveId: string;
  itemId?: string;
  top?: number;
}

export async function listLibraryItems(params: ListLibraryItemsParams) {
  const token = await getToken();
  const { driveId, itemId, top = 20 } = params;

  const basePath = itemId
    ? `/drives/${driveId}/items/${itemId}/children`
    : `/drives/${driveId}/root/children`;

  const queryParams = new URLSearchParams({
    $top: String(top),
    $select: "id,name,size,lastModifiedDateTime,folder,file,webUrl",
  });

  const result = await graphFetch(token, `${basePath}?${queryParams}`);
  return result.value;
}

export async function searchSharePoint(query: string) {
  const token = await getToken();
  const result = await graphFetch(token, "/search/query", {
    method: "POST",
    body: JSON.stringify({
      requests: [
        {
          entityTypes: ["driveItem", "listItem", "site"],
          query: { queryString: query },
          from: 0,
          size: 10,
        },
      ],
    }),
  });
  return result.value;
}
