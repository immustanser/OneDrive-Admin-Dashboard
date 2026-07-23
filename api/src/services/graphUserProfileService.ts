import { getGraphAccessToken } from './graphAuthService';
import { fetchWithRetry } from '../utils/graphFetch';

const GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';

export interface IUserProfileInfo {
  department: string;
  jobTitle: string;
  manager: string;
}

const EMPTY_PROFILE: IUserProfileInfo = { department: '', jobTitle: '', manager: '' };

/**
 * In-memory cache of user profile lookups, keyed by userPrincipalName.
 * Scoped at module level (like the token cache in graphAuthService.ts) so
 * it persists for the lifetime of the warm Function host.
 *
 * IMPORTANT: this service is intentionally called PER-USER, on demand
 * (see functions/getUserProfile.ts), not for the full tenant. On large
 * tenants (thousands of OneDrive owners), calling /users/{upn} and
 * /users/{upn}/manager for every inventory row causes Microsoft Graph to
 * respond with 429 Too Many Requests. The inventory report endpoints
 * (getOneDriveDashboardData / getOneDriveAccountDetails) only use the
 * bulk usage reports and never call this service in a loop; the SPFx
 * client instead calls GET /api/user-profile?upn=... only for the rows
 * currently visible on-screen (see usePagedOneDriveUsers.ts).
 */
const profileCache = new Map<string, Promise<IUserProfileInfo>>();

async function fetchUserDepartmentAndTitle(
  userPrincipalName: string,
  token: string
): Promise<{ department: string; jobTitle: string }> {
  const url = `${GRAPH_BASE_URL}/users/${encodeURIComponent(userPrincipalName)}?$select=department,jobTitle`;
  const response = await fetchWithRetry(
    url,
    { method: 'GET', headers: { Authorization: `Bearer ${token}` } },
    `user profile ${userPrincipalName}`
  );

  if (!response.ok) {
    // Missing/disabled users, or lookup failures, should not break the
    // whole dashboard — fall back to empty values for this user.
    return { department: '', jobTitle: '' };
  }

  const body = (await response.json()) as { department?: string; jobTitle?: string };
  return {
    department: body.department || '',
    jobTitle: body.jobTitle || ''
  };
}

async function fetchUserManagerDisplayName(userPrincipalName: string, token: string): Promise<string> {
  const url = `${GRAPH_BASE_URL}/users/${encodeURIComponent(userPrincipalName)}/manager?$select=displayName`;
  const response = await fetchWithRetry(
    url,
    { method: 'GET', headers: { Authorization: `Bearer ${token}` } },
    `user manager ${userPrincipalName}`
  );

  if (!response.ok) {
    // 404 is expected for users with no manager assigned; any other
    // failure is also treated as "no manager" rather than failing the
    // whole lookup.
    return '';
  }

  const body = (await response.json()) as { displayName?: string };
  return body.displayName || '';
}

async function fetchUserProfile(userPrincipalName: string): Promise<IUserProfileInfo> {
  const token = await getGraphAccessToken();

  const [{ department, jobTitle }, manager] = await Promise.all([
    fetchUserDepartmentAndTitle(userPrincipalName, token),
    fetchUserManagerDisplayName(userPrincipalName, token)
  ]);

  return { department, jobTitle, manager };
}

/**
 * Retrieves department, job title and manager display name for a single
 * user via Microsoft Graph (GET /users/{userPrincipalName} and
 * GET /users/{userPrincipalName}/manager), using the same app-only token as
 * the report calls. Results are cached in-memory per userPrincipalName so
 * repeated lookups (e.g. re-visiting the same inventory page) never
 * trigger a duplicate Graph call.
 */
export async function getUserProfile(userPrincipalName: string): Promise<IUserProfileInfo> {
  if (!userPrincipalName) {
    return EMPTY_PROFILE;
  }

  const key = userPrincipalName.trim().toLowerCase();
  let pending = profileCache.get(key);
  if (!pending) {
    pending = fetchUserProfile(userPrincipalName).catch(() => EMPTY_PROFILE);
    profileCache.set(key, pending);
  }
  return pending;
}
