import { getGraphAccessToken } from './graphAuthService';
import { fetchWithRetry } from '../utils/graphFetch';

const GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';

/**
 * In-flight de-duplication: if a report is already being downloaded
 * (including its 429 retry loop), any other caller for the SAME report
 * path reuses that same promise instead of starting a second, parallel
 * download/retry loop against Microsoft Graph. This is on top of (and
 * independent from) the higher-level 30-minute dashboard cache in
 * dashboardCacheService.ts, which is what normally prevents this
 * function from being called at all on a warm cache hit.
 */
const pendingReportFetches = new Map<string, Promise<string>>();

// Microsoft Graph "usage report" endpoints normally respond with a 302
// redirect whose `Location` header points to a short-lived,
// pre-authenticated SAS URL that serves the actual CSV content.
//
// The 302 redirect from a Graph usage report endpoint points to a
// short-lived, pre-authenticated SAS URL. That URL can occasionally come
// back with a 403 (e.g. if it takes a moment to become valid, or if it
// is treated as effectively single-use) even though the report request
// itself succeeded. Re-requesting the report endpoint yields a brand new
// redirect URL, so a small number of retries here resolves this without
// masking a real permissions problem (which would fail on the initial
// Graph call, not on the redirect download).
const MAX_REDIRECT_DOWNLOAD_ATTEMPTS = 3;

async function requestReportRedirect(reportPath: string): Promise<string> {
  const token = await getGraphAccessToken();
  const url = `${GRAPH_BASE_URL}${reportPath}`;

  const response = await fetchWithRetry(
    url,
    {
      method: 'GET',
      redirect: 'manual',
      headers: {
        Authorization: `Bearer ${token}`
      }
    },
    `report ${reportPath}`
  );

  if (response.status === 302 || response.status === 301) {
    const location = response.headers.get('location');
    if (!location) {
      throw new Error(`Graph report redirect (${response.status}) had no Location header for ${reportPath}.`);
    }
    return location;
  }

  if (response.ok) {
    // Some report responses (or a proxy in front of Graph) may already
    // return the CSV body directly with a 200, without a redirect.
    return response.text().then((body) => `data:text/csv,${encodeURIComponent(body)}`);
  }

  const errorBody = await response.text().catch(() => '');
  throw new Error(`Graph report request failed for ${reportPath} (status ${response.status}): ${errorBody}`);
}

async function getReportCsv(reportPath: string): Promise<string> {
  const existing = pendingReportFetches.get(reportPath);
  if (existing) {
    // eslint-disable-next-line no-console
    console.log(`Report fetch already in progress for ${reportPath}, reusing existing promise.`);
    return existing;
  }

  const fetchPromise = (async (): Promise<string> => {
    for (let attempt = 1; attempt <= MAX_REDIRECT_DOWNLOAD_ATTEMPTS; attempt++) {
      const location = await requestReportRedirect(reportPath);

      if (location.startsWith('data:text/csv,')) {
        return decodeURIComponent(location.substring('data:text/csv,'.length));
      }

      const csvResponse = await fetchWithRetry(location, { method: 'GET' }, `report ${reportPath} (redirect download)`);
      if (csvResponse.ok) {
        return await csvResponse.text();
      }

      if (csvResponse.status === 403 && attempt < MAX_REDIRECT_DOWNLOAD_ATTEMPTS) {
        // eslint-disable-next-line no-console
        console.warn(
          `Redirect download for ${reportPath} returned 403 (attempt ${attempt}/${MAX_REDIRECT_DOWNLOAD_ATTEMPTS}). ` +
          `Requesting a fresh report redirect and retrying.`
        );
        continue;
      }

      throw new Error(`Failed downloading CSV report from redirect location (status ${csvResponse.status}).`);
    }

    throw new Error(`Failed downloading CSV report for ${reportPath} after ${MAX_REDIRECT_DOWNLOAD_ATTEMPTS} attempts.`);
  })();

  pendingReportFetches.set(reportPath, fetchPromise);
  try {
    return await fetchPromise;
  } finally {
    pendingReportFetches.delete(reportPath);
  }
}

export async function getOneDriveUsageAccountDetailCsv(period: string = 'D180'): Promise<string> {
  return getReportCsv(`/reports/getOneDriveUsageAccountDetail(period='${period}')`);
}

export async function getOneDriveUsageStorageCsv(period: string = 'D180'): Promise<string> {
  return getReportCsv(`/reports/getOneDriveUsageStorage(period='${period}')`);
}
