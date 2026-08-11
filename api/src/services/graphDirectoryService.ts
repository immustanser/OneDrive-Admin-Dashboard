import { getGraphAccessToken } from './graphAuthService';
import { fetchWithRetry } from '../utils/graphFetch';

const GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';

interface IGraphUserDirectoryPage {
  value: { userPrincipalName?: string; department?: string }[];
  '@odata.nextLink'?: string;
}

/**
 * Fetches department for every user in the tenant using Microsoft
 * Graph's paged /users listing (~9 requests for ~9,000 users at $top=999),
 * NOT one request per OneDrive owner. This is fundamentally different
 * from the per-user throttling problem the rest of this backend avoids
 * (see dashboardMapperService.ts / graphUserProfileService.ts) - it is a
 * small, bounded number of bulk/paged requests, reuses the already
 * approved Directory.Read.All application permission, and its result is
 * cached alongside the rest of the dashboard payload by
 * dashboardCacheService.ts, so it only runs once per cache refresh
 * (every 30 minutes, or on demand via "Refresh Data") rather than once
 * per dashboard load.
 *
 * Returns a map of lower-cased userPrincipalName -> department (users
 * with no department set are simply omitted).
 */
export async function getAllUserDepartments(): Promise<Map<string, string>> {
  const token = await getGraphAccessToken();
  const departmentsByUpn = new Map<string, string>();

  let url: string | undefined =
    `${GRAPH_BASE_URL}/users?$select=userPrincipalName,department&$top=999`;

  while (url) {
    const response = await fetchWithRetry(
      url,
      {
        method: 'GET',
        headers: { Authorization: `Bearer ${token}` }
      },
      'GET /users (bulk department directory)'
    );

    if (!response.ok) {
      throw new Error(
        `Microsoft Graph /users request failed with status ${response.status} ${response.statusText}.`
      );
    }

    const page = (await response.json()) as IGraphUserDirectoryPage;
    page.value.forEach((u) => {
      if (u.userPrincipalName && u.department && u.department.trim()) {
        departmentsByUpn.set(u.userPrincipalName.toLowerCase(), u.department.trim());
      }
    });

    url = page['@odata.nextLink'];
  }

  return departmentsByUpn;
}
