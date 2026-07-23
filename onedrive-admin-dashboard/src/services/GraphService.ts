import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IDashboardKpis, IOneDriveUser, ISharingReport, IGovernanceRiskItem, IStorageTrendPoint, ITopOneDrive } from '../models';

/**
 * SECURITY NOTE
 * -------------
 * This service NEVER calls Microsoft Graph directly and NEVER holds a
 * client secret. All tenant OneDrive usage data is retrieved from the
 * secure Azure Function backend (see /api), which performs the
 * Microsoft Graph client-credentials authentication server-side.
 * Only the Azure Function's public HTTPS endpoint URL is configured here
 * (via the web part's "Azure Function API Base URL" property).
 */

export interface IInactiveBuckets {
  inactive30Days: number;
  inactive60Days: number;
  inactive90Days: number;
}

export interface IStorageAnalytics {
  storageTrend: IStorageTrendPoint[];
  topOneDrives: ITopOneDrive[];
  storageByDepartment: { department: string; storageUsedGB: number }[];
}

export interface IOneDriveDashboardApiResponse {
  kpiData: IDashboardKpis;
  inventoryUsers: IOneDriveUser[];
  storageAnalytics: IStorageAnalytics;
  // Sharing report is no longer rendered anywhere in the UI. Kept as an
  // optional field for backward compatibility with the API response -
  // see components/sharing (unused) and models/ISharingReport.ts.
  sharingReportPlaceholders?: ISharingReport;
  inactiveOneDrives: IInactiveBuckets;
  governanceRisks: IGovernanceRiskItem[];
  generatedDateTime: string;
  dataSource: string;
}

export interface IUserProfileApiResponse {
  department: string;
  jobTitle: string;
  manager: string;
}

const DASHBOARD_ROUTE = '/api/onedrive-dashboard';
const USER_PROFILE_ROUTE = '/api/user-profile';

export class GraphService {
  private static _apiBaseUrl: string | undefined;

  /**
   * The WebPartContext parameter is accepted for API symmetry with other
   * services (and potential future use, e.g. telemetry) even though the
   * current implementation only needs the Azure Function base URL.
   */
  public static init(context: WebPartContext, apiBaseUrl: string): void {
    this._apiBaseUrl = (apiBaseUrl || '').trim().replace(/\/+$/, '');
  }

  /**
   * Calls the secure Azure Function backend to retrieve the full OneDrive
   * dashboard payload. The Function itself authenticates to Microsoft
   * Graph using the client credentials flow with a secret that is only
   * ever stored in Azure Function App Settings / Key Vault.
   */
  public static async getDashboardData(forceRefresh: boolean = false): Promise<IOneDriveDashboardApiResponse> {
    if (!this._apiBaseUrl) {
      throw new Error(
        'The Azure Function API Base URL has not been configured. Set it in the web part property pane.'
      );
    }

    const url = `${this._apiBaseUrl}${DASHBOARD_ROUTE}${forceRefresh ? '?forceRefresh=true' : ''}`;
    let response: Response;
    try {
      // eslint-disable-next-line no-console
      console.log('GraphService: requesting OneDrive dashboard data from', url);
      response = await fetch(url, { method: 'GET', headers: { Accept: 'application/json' } });
    } catch (err) {
      // eslint-disable-next-line no-console
      console.error('GraphService: Dashboard API load FAILED (network error).', err);
      throw new Error(
        'Unable to reach the OneDrive Dashboard data service. Please verify the Azure Function endpoint or contact your administrator.'
      );
    }

    if (!response.ok) {
      let message = `Request to the OneDrive Dashboard data service failed (status ${response.status}).`;
      try {
        const body = await response.json();
        // The backend returns a friendly, non-retryable message with
        // isThrottled=true when Microsoft Graph is throttling report
        // downloads (HTTP 429) - surface that message as-is so the UI
        // shows guidance instead of a generic error.
        if (body && (body.message || body.error)) {
          message = body.message || body.error;
        }
      } catch {
        // ignore body parse failures, use default message
      }
      // eslint-disable-next-line no-console
      console.error('GraphService: Dashboard API load FAILED.', message);
      throw new Error(message);
    }

    const data = (await response.json()) as IOneDriveDashboardApiResponse;
    // eslint-disable-next-line no-console
    console.log('Dashboard API loaded successfully');
    // eslint-disable-next-line no-console
    console.log('GraphService: dataSource =', data.dataSource, '| generatedDateTime =', data.generatedDateTime);
    return data;
  }

  /**
   * Calls the secure Azure Function backend to retrieve profile info
   * (department, jobTitle, manager) for a SINGLE user. Intended to be
   * called only for rows currently visible in the Inventory grid, never
   * for the full tenant, to avoid Microsoft Graph throttling.
   */
  public static async getUserProfile(userPrincipalName: string): Promise<IUserProfileApiResponse> {
    if (!this._apiBaseUrl) {
      throw new Error(
        'The Azure Function API Base URL has not been configured. Set it in the web part property pane.'
      );
    }

    const url = `${this._apiBaseUrl}${USER_PROFILE_ROUTE}?upn=${encodeURIComponent(userPrincipalName)}`;
    const response = await fetch(url, { method: 'GET', headers: { Accept: 'application/json' } });

    if (!response.ok) {
      throw new Error(`Request to the user profile service failed (status ${response.status}) for ${userPrincipalName}.`);
    }

    return (await response.json()) as IUserProfileApiResponse;
  }
}
