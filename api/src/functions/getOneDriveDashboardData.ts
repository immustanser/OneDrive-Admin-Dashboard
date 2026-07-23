import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getOneDriveUsageAccountDetailCsv, getOneDriveUsageStorageCsv } from '../services/graphReportsService';
import { buildDashboardResponse } from '../services/dashboardMapperService';
import { getOrBuildDashboard } from '../services/dashboardCacheService';
import { GraphThrottledError } from '../utils/graphFetch';

/**
 * GET /api/onedrive-dashboard[?forceRefresh=true]
 *
 * Combined endpoint used by the SPFx web part. Acquires an app-only Graph
 * token via client credentials (see graphAuthService.ts), downloads the
 * OneDrive usage CSV reports (bulk reports ONLY - no per-user profile
 * calls happen here), maps them to dashboard-ready JSON, and returns a
 * single payload. No tokens or secrets are ever included in the response.
 *
 * Response is served from a 30-minute in-memory backend cache
 * (dashboardCacheService.ts) unless `forceRefresh=true` is passed - the
 * SPFx "Refresh Data" button is the only caller expected to pass that.
 * Concurrent requests (e.g. multiple tabs/users) share a single in-flight
 * build instead of each triggering their own Microsoft Graph Reports API
 * calls.
 *
 * NOTE ON AUTH LEVEL: this function is registered with authLevel
 * "anonymous" for simplicity of local development and initial wiring.
 * Before production use, lock this down (Azure AD / Easy Auth, APIM,
 * a function key, and/or restrict CORS to the SharePoint tenant origin).
 */
export async function getOneDriveDashboardData(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  try {
    const forceRefresh = request.query.get('forceRefresh') === 'true';

    const dashboard = await getOrBuildDashboard(forceRefresh, async () => {
      // Fetched sequentially (not Promise.all) - Microsoft Graph's usage
      // Reports API throttles much more aggressively when the same app
      // issues concurrent report requests. Downloading one report at a
      // time meaningfully reduces 429s without changing the resulting
      // data or the dashboard response shape.
      const accountDetailCsv = await getOneDriveUsageAccountDetailCsv('D180');
      const storageCsv = await getOneDriveUsageStorageCsv('D180');
      return buildDashboardResponse(accountDetailCsv, storageCsv);
    });

    return {
      status: 200,
      jsonBody: dashboard
    };
  } catch (error) {
    if (error instanceof GraphThrottledError) {
      context.warn('getOneDriveDashboardData: Microsoft Graph is throttling report downloads.', error);
      return {
        status: 429,
        jsonBody: {
          success: false,
          message: 'Microsoft Graph is throttling report downloads. Please try Refresh Data later.',
          isThrottled: true
        }
      };
    }

    context.error('getOneDriveDashboardData failed:', error);
    return {
      status: 502,
      jsonBody: {
        error: 'Unable to retrieve OneDrive dashboard data from Microsoft Graph.',
        message: error instanceof Error ? error.message : 'Unknown error'
      }
    };
  }
}

app.http('getOneDriveDashboardData', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'onedrive-dashboard',
  handler: getOneDriveDashboardData
});
