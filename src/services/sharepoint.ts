import type { AuthProvider } from "../types/auth.js";
import type {
  GraphSite,
  GraphDrive,
  GraphDriveItem,
  GraphSearchResponse,
  GraphSearchHit,
  GraphPagedResponse,
} from "../types/graph.js";
import { graphFetch } from "../utils/graph-client.js";
import { SCOPES, DEFAULT_PAGE_SIZE_SMALL, DEFAULT_PAGE_SIZE_LARGE } from "../constants.js";
import { createGetToken } from "../utils/auth-helper.js";
import { flagName } from "../utils/flagged-paths.js";

function flagSearchHit(hit: GraphSearchHit): GraphSearchHit {
  const decodedUrl = hit.resource.webUrl ? decodeURIComponent(hit.resource.webUrl) : undefined;
  return {
    ...hit,
    resource: {
      ...hit.resource,
      name: hit.resource.name ? flagName(hit.resource.name, decodedUrl) : hit.resource.name,
      displayName: hit.resource.displayName
        ? flagName(hit.resource.displayName, decodedUrl)
        : hit.resource.displayName,
    },
  };
}

export interface ListLibraryItemsParams {
  driveId: string;
  itemId?: string;
  top?: number;
}

export function createSharePointService(auth: AuthProvider) {
  const getToken = createGetToken(auth, SCOPES.SHAREPOINT);

  async function listSites(query?: string): Promise<GraphSite[]> {
    const token = await getToken();
    const searchQuery = query ?? "*";
    const result = await graphFetch<GraphPagedResponse<GraphSite>>(
      token,
      `/sites?search=${encodeURIComponent(searchQuery)}&$select=id,displayName,name,webUrl,description`
    );
    return result.value;
  }

  async function getSite(siteId: string): Promise<GraphSite> {
    const token = await getToken();
    return graphFetch<GraphSite>(
      token,
      `/sites/${siteId}?$select=id,displayName,name,webUrl,description`
    );
  }

  async function listDocumentLibraries(siteId: string): Promise<GraphDrive[]> {
    const token = await getToken();
    const result = await graphFetch<GraphPagedResponse<GraphDrive>>(
      token,
      `/sites/${siteId}/drives?$select=id,name,driveType,webUrl,description`
    );
    return result.value;
  }

  async function listLibraryItems(params: ListLibraryItemsParams): Promise<GraphDriveItem[]> {
    const token = await getToken();
    const { driveId, itemId, top = DEFAULT_PAGE_SIZE_LARGE } = params;

    const basePath = itemId
      ? `/drives/${driveId}/items/${itemId}/children`
      : `/drives/${driveId}/root/children`;

    const queryParams = new URLSearchParams({
      $top: String(top),
      $select: "id,name,size,lastModifiedDateTime,folder,file,webUrl,parentReference",
    });

    const result = await graphFetch<GraphPagedResponse<GraphDriveItem>>(
      token,
      `${basePath}?${queryParams}`
    );
    return result.value.map((item) => ({
      ...item,
      name: flagName(item.name, item.parentReference?.path),
    }));
  }

  async function searchSharePoint(query: string): Promise<GraphSearchResponse[]> {
    const token = await getToken();
    const result = await graphFetch<GraphPagedResponse<GraphSearchResponse>>(
      token,
      "/search/query",
      {
        method: "POST",
        body: JSON.stringify({
          requests: [
            {
              entityTypes: ["driveItem", "listItem", "site"],
              query: { queryString: query },
              from: 0,
              size: DEFAULT_PAGE_SIZE_SMALL,
            },
          ],
        }),
      }
    );
    return result.value.map((response) => ({
      ...response,
      hitsContainers: response.hitsContainers.map((container) => ({
        ...container,
        hits: container.hits.map(flagSearchHit),
      })),
    }));
  }

  return { listSites, getSite, listDocumentLibraries, listLibraryItems, searchSharePoint };
}

export type SharePointService = ReturnType<typeof createSharePointService>;
