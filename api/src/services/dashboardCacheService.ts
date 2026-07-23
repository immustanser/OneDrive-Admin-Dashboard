import { IOneDriveDashboardResponse } from '../models/IOneDriveDashboardResponse';

/**
 * In-memory backend cache for the full OneDrive dashboard payload.
 *
 * Scope/lifetime: module-level, so it persists for the lifetime of the
 * warm Azure Functions host (same pattern as the Graph token cache in
 * graphAuthService.ts). This is what prevents every SPFx page
 * load/refresh from re-downloading the (large) Microsoft Graph usage
 * reports and re-triggering 429 throttling.
 *
 * Cache key: 'onedrive-dashboard-D180' (the only report period currently
 * used by the dashboard). Cache duration: 30 minutes. A build already in
 * progress is also de-duplicated (getOrBuildDashboard) so that concurrent
 * requests arriving before the first completes never trigger a second,
 * parallel set of Graph Reports API calls.
 */

const CACHE_KEY = 'onedrive-dashboard-D180';
const CACHE_TTL_MS = 30 * 60 * 1000;

interface ICacheEntry {
  data: IOneDriveDashboardResponse;
  expiresAt: number;
}

let cached: ICacheEntry | undefined;
let inFlight: Promise<IOneDriveDashboardResponse> | undefined;

function getCached(): IOneDriveDashboardResponse | undefined {
  if (cached && cached.expiresAt > Date.now()) {
    return cached.data;
  }
  return undefined;
}

function setCached(data: IOneDriveDashboardResponse): void {
  cached = { data, expiresAt: Date.now() + CACHE_TTL_MS };
}

/**
 * Clears the cached dashboard payload. Called only when a caller
 * explicitly requests a forced refresh (the "Refresh Data" button).
 */
export function invalidateDashboardCache(): void {
  cached = undefined;
}

/**
 * Returns the cached dashboard payload if present and not force-refreshed,
 * otherwise builds a fresh one via `builder()`. Concurrent calls while a
 * build is already in progress share that SAME in-flight promise instead
 * of starting independent Microsoft Graph Reports API calls.
 */
export async function getOrBuildDashboard(
  forceRefresh: boolean,
  builder: () => Promise<IOneDriveDashboardResponse>
): Promise<IOneDriveDashboardResponse> {
  if (forceRefresh) {
    cached = undefined;
  } else {
    const hit = getCached();
    if (hit) {
      // eslint-disable-next-line no-console
      console.log(`Dashboard response served from backend cache (key=${CACHE_KEY}).`);
      return hit;
    }
  }

  if (inFlight) {
    // eslint-disable-next-line no-console
    console.log('Dashboard build already in progress, reusing existing promise.');
    return inFlight;
  }

  // eslint-disable-next-line no-console
  console.log(`Dashboard cache miss, calling Microsoft Graph Reports API (key=${CACHE_KEY}).`);

  inFlight = builder()
    .then((data) => {
      setCached(data);
      // eslint-disable-next-line no-console
      console.log('Dashboard cache refreshed successfully.');
      return data;
    })
    .finally(() => {
      inFlight = undefined;
    });

  return inFlight;
}
