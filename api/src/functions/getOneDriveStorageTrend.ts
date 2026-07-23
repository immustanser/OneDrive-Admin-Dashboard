import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getOneDriveUsageStorageCsv } from '../services/graphReportsService';
import { parseStorageCsv, computeStorageTrend } from '../services/dashboardMapperService';

/**
 * GET /api/onedrive-storage-trend
 *
 * Standalone endpoint returning just the monthly storage trend (mapped
 * from getOneDriveUsageStorage). Useful for refreshing the Storage
 * Analytics tab independently of the full dashboard payload.
 */
export async function getOneDriveStorageTrend(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  try {
    const csv = await getOneDriveUsageStorageCsv('D180');
    const storageTrend = computeStorageTrend(parseStorageCsv(csv));

    return {
      status: 200,
      jsonBody: { storageTrend }
    };
  } catch (error) {
    context.error('getOneDriveStorageTrend failed:', error);
    return {
      status: 502,
      jsonBody: {
        error: 'Unable to retrieve OneDrive storage trend from Microsoft Graph.',
        message: error instanceof Error ? error.message : 'Unknown error'
      }
    };
  }
}

app.http('getOneDriveStorageTrend', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'onedrive-storage-trend',
  handler: getOneDriveStorageTrend
});
