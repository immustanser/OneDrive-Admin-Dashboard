import { WebPartContext } from '@microsoft/sp-webpart-base';
import { AadHttpClient, HttpClientResponse } from '@microsoft/sp-http';
import { IDashboardKpis, IOneDriveUser, ISharingReport, IGovernanceRiskItem, IStorageTrendPoint, ITopOneDrive } from '../models';

/**
 * SECURITY NOTE
 * -------------
 * This service NEVER calls Microsoft Graph directly and NEVER holds a
 * client secret. All tenant OneDrive usage data is retrieved from the
 * secure Azure Function backend (see /api), which performs the
 * Microsoft Graph client-credentials authentication server-side.
 *
 * The Azure Function itself is secured with Microsoft Entra ID
 * (App Service Authentication) and rejects unauthenticated requests with
 * HTTP 401. This service authenticates to it using SPFx's built-in
 * AadHttpClient (via AadHttpClientFactory), which silently acquires and
 * attaches an Azure AD access token (Authorization: Bearer ...) for the
 * signed-in SharePoint user - no function key and no client secret is
 * ever present in this code or in the browser.
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

const AUTH_FAILURE_MESSAGE =
  'Unable to authenticate to the OneDrive Dashboard API. Please verify that API permissions have been approved in SharePoint Admin Center.';
const UNAUTHORIZED_MESSAGE =
  'Access to the OneDrive Dashboard API is unauthorized. Please verify Microsoft Entra authentication and SPFx API permissions.';
const FORBIDDEN_MESSAGE =
  'You are authenticated but not authorized to access the OneDrive Dashboard API.';

export class GraphService {
  private static _apiBaseUrl: string | undefined;
  private static _apiResourceUri: string | undefined;
  private static _context: WebPartContext | undefined;
  private static _aadClientPromise: Promise<AadHttpClient> | undefined;

  /**
   * The WebPartContext is used to obtain an AadHttpClient (via
   * context.aadHttpClientFactory) that automatically acquires and
   * attaches a Microsoft Entra ID access token to every request made to
   * the secured Azure Function. apiResourceUri must be the App ID URI
   * (or exposed API resource) configured on the Function App's Entra ID
   * authentication app registration - see README for exact steps.
   */
  public static init(context: WebPartContext, apiBaseUrl: string, apiResourceUri: string): void {
    this._apiBaseUrl = (apiBaseUrl || '').trim().replace(/\/+$/, '');
    this._apiResourceUri = (apiResourceUri || '').trim();
    this._context = context;
    // Reset any previously cached client - a new resource/context means
    // a new client must be requested from the factory.
    this._aadClientPromise = undefined;
  }

  /**
   * Lazily obtains (and caches) the AadHttpClient for the secured Azure
   * Function resource. All calls to the Function go through this client
   * instead of plain fetch, so every request carries an
   * "Authorization: Bearer <token>" header - never a function key, never
   * a client secret.
   */
  private static getClient(): Promise<AadHttpClient> {
    if (!this._context) {
      return Promise.reject(new Error(AUTH_FAILURE_MESSAGE));
    }
    if (!this._apiResourceUri) {
      return Promise.reject(new Error(
        'The Azure Function API Resource (App ID URI) has not been configured. Set it in the web part property pane.'
      ));
    }
    if (!this._aadClientPromise) {
      // eslint-disable-next-line no-console
      console.log('Using AadHttpClient for secured Azure Function API');
      this._aadClientPromise = this._context.aadHttpClientFactory
        .getClient(this._apiResourceUri)
        .catch((err) => {
          // Allow a later call to retry instead of being stuck reusing a
          // rejected promise forever.
          this._aadClientPromise = undefined;
          // eslint-disable-next-line no-console
          console.error('GraphService: AadHttpClient token acquisition FAILED.', err);
          throw new Error(AUTH_FAILURE_MESSAGE);
        });
    }
    return this._aadClientPromise;
  }

  /**
   * Translates a non-ok AadHttpClient response into a friendly error
   * message, without ever surfacing raw token/secret details.
   */
  private static async toFriendlyError(response: HttpClientResponse, context: string): Promise<Error> {
    if (response.status === 401) {
      // eslint-disable-next-line no-console
      console.error(`GraphService: ${context} returned 401`);
      return new Error(UNAUTHORIZED_MESSAGE);
    }
    if (response.status === 403) {
      // eslint-disable-next-line no-console
      console.error(`GraphService: ${context} returned 403`);
      return new Error(FORBIDDEN_MESSAGE);
    }

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
    console.error(`GraphService: ${context} FAILED.`, message);
    return new Error(message);
  }

  /**
   * Calls the secure Azure Function backend to retrieve the full OneDrive
   * dashboard payload. The Function itself authenticates to Microsoft
   * Graph using the client credentials flow with a secret that is only
   * ever stored in Azure Function App Settings / Key Vault. This request
   * is authenticated with an Entra ID token via AadHttpClient.
   */
  public static async getDashboardData(forceRefresh: boolean = false): Promise<IOneDriveDashboardApiResponse> {
    if (!this._apiBaseUrl) {
      throw new Error(
        'The Azure Function API Base URL has not been configured. Set it in the web part property pane.'
      );
    }

    const url = `${this._apiBaseUrl}${DASHBOARD_ROUTE}${forceRefresh ? '?forceRefresh=true' : ''}`;
    const client = await this.getClient();

    // eslint-disable-next-line no-console
    console.log('Dashboard API authenticated request started');
    // eslint-disable-next-line no-console
    console.log('GraphService: requesting OneDrive dashboard data from', url);

    let response: HttpClientResponse;
    try {
      response = await client.get(url, AadHttpClient.configurations.v1);
    } catch (err) {
      // eslint-disable-next-line no-console
      console.error('GraphService: Dashboard API load FAILED (network error).', err);
      throw new Error(
        'Unable to reach the OneDrive Dashboard data service. Please verify the Azure Function endpoint or contact your administrator.'
      );
    }

    if (!response.ok) {
      throw await this.toFriendlyError(response, 'Dashboard API authenticated request');
    }

    const data = (await response.json()) as IOneDriveDashboardApiResponse;
    // eslint-disable-next-line no-console
    console.log('Dashboard API authenticated request succeeded');
    // eslint-disable-next-line no-console
    console.log('GraphService: dataSource =', data.dataSource, '| generatedDateTime =', data.generatedDateTime);
    return data;
  }

  /**
   * Calls the secure Azure Function backend to retrieve profile info
   * (department, jobTitle, manager) for a SINGLE user. Intended to be
   * called only for rows currently visible in the Inventory grid, never
   * for the full tenant, to avoid Microsoft Graph throttling. This
   * request is authenticated with an Entra ID token via AadHttpClient.
   */
  public static async getUserProfile(userPrincipalName: string): Promise<IUserProfileApiResponse> {
    if (!this._apiBaseUrl) {
      throw new Error(
        'The Azure Function API Base URL has not been configured. Set it in the web part property pane.'
      );
    }

    const url = `${this._apiBaseUrl}${USER_PROFILE_ROUTE}?upn=${encodeURIComponent(userPrincipalName)}`;
    const client = await this.getClient();

    // eslint-disable-next-line no-console
    console.log('User profile authenticated request started');

    const response = await client.get(url, AadHttpClient.configurations.v1);

    if (!response.ok) {
      throw await this.toFriendlyError(response, `User profile authenticated request for ${userPrincipalName}`);
    }

    return (await response.json()) as IUserProfileApiResponse;
  }
}
