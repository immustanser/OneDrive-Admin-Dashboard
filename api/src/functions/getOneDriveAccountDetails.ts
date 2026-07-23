import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getOneDriveUsageAccountDetailCsv } from '../services/graphReportsService';
import { parseAccountDetailCsv, mapToInventoryUsers } from '../services/dashboardMapperService';
import { GraphThrottledError } from '../utils/graphFetch';

/**
 * GET /api/onedrive-account-details
 *
 * Standalone endpoint returning just the OneDrive inventory (mapped from
 * getOneDriveUsageAccountDetail). Useful for refreshing the Inventory tab
 * independently of the full dashboard payload.
 */
export async function getOneDriveAccountDetails(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  try {
    const csv = await getOneDriveUsageAccountDetailCsv('D180');
    const users = mapToInventoryUsers(parseAccountDetailCsv(csv));

    return {
      status: 200,
      jsonBody: { inventoryUsers: users }
    };
  } catch (error) {
    if (error instanceof GraphThrottledError) {
      context.error('getOneDriveAccountDetails throttled by Microsoft Graph:', error);
      return {
        status: 429,
        jsonBody: {
          success: false,
          message: 'Microsoft Graph is throttling report downloads. Please try Refresh Data later.',
          isThrottled: true
        }
      };
    }
    context.error('getOneDriveAccountDetails failed:', error);
    return {
      status: 502,
      jsonBody: {
        error: 'Unable to retrieve OneDrive account details from Microsoft Graph.',
        message: error instanceof Error ? error.message : 'Unknown error'
      }
    };
  }
}

app.http('getOneDriveAccountDetails', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'onedrive-account-details',
  handler: getOneDriveAccountDetails
});
