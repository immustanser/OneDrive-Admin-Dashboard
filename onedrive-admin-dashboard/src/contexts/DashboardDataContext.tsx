import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { OneDriveService } from '../services';
import { IDashboardKpis, IOneDriveUser, IGovernanceRiskItem, IStorageTrendPoint, ITopOneDrive } from '../models';
import { ReportService } from '../services/ReportService';

export interface IDashboardDataState {
  loading: boolean;
  error: string | undefined;
  kpis: IDashboardKpis | undefined;
  users: IOneDriveUser[];
  risks: IGovernanceRiskItem[];
  storageTrend: IStorageTrendPoint[];
  topOneDrives: ITopOneDrive[];
  refresh: () => void;
  /**
   * Root URL of the tenant's SharePoint site (e.g. "https://contoso.sharepoint.com"),
   * derived from the SPFx WebPartContext. Used to build real links (e.g. user
   * profile pages, OneDrive personal-site fallback URLs) instead of hardcoded
   * placeholder domains. Empty string if it cannot be determined.
   */
  tenantRootUrl: string;
}

const DashboardDataContext = React.createContext<IDashboardDataState | undefined>(undefined);

export const DashboardDataProvider: React.FC<{
  context: WebPartContext;
  apiBaseUrl: string;
  apiResourceUri: string;
  useMockData: boolean;
  children: React.ReactNode;
}> = ({
  context,
  apiBaseUrl,
  apiResourceUri,
  useMockData,
  children
}) => {
  const [loading, setLoading] = React.useState<boolean>(true);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const [kpis, setKpis] = React.useState<IDashboardKpis | undefined>(undefined);
  const [users, setUsers] = React.useState<IOneDriveUser[]>([]);
  const [risks, setRisks] = React.useState<IGovernanceRiskItem[]>([]);
  const [storageTrend, setStorageTrend] = React.useState<IStorageTrendPoint[]>([]);
  const [topOneDrives, setTopOneDrives] = React.useState<ITopOneDrive[]>([]);
  const [refreshToken, setRefreshToken] = React.useState<number>(0);

  const tenantRootUrl = React.useMemo<string>(() => {
    try {
      const absoluteUrl = context && context.pageContext && context.pageContext.web
        ? context.pageContext.web.absoluteUrl
        : '';
      return absoluteUrl ? new URL(absoluteUrl).origin : '';
    } catch {
      return '';
    }
  }, [context]);

  React.useEffect(() => {
    let isMounted = true;
    OneDriveService.init(context, apiBaseUrl, apiResourceUri, useMockData);

    async function load(): Promise<void> {
      setLoading(true);
      setError(undefined);
      try {
        // Single source of truth: the dashboard payload is fetched
        // EXACTLY ONCE here (OneDriveService.getDashboardSnapshot() is
        // internally deduplicated + 10-minute cached, so even if this
        // effect re-runs it will not cause a second network request).
        // All tabs/components consume the resulting state via
        // useDashboardData() - none of them call the API independently.
        const snapshot = await OneDriveService.getDashboardSnapshot();

        let kpiData: IDashboardKpis | undefined;
        let userData: IOneDriveUser[];
        let riskData: IGovernanceRiskItem[];
        let trend: IStorageTrendPoint[];
        let top: ITopOneDrive[];

        if (snapshot) {
          // Live data path: everything comes from the single dashboard
          // response - no additional service/API calls.
          kpiData = snapshot.kpiData;
          userData = snapshot.inventoryUsers;
          riskData = snapshot.governanceRisks;
          trend = snapshot.storageAnalytics.storageTrend;
          top = snapshot.storageAnalytics.topOneDrives.slice(0, 10);
        } else {
          // Mock-data path (local development only): no Azure Function
          // calls are involved here at all, so there is no duplicate
          // request concern - each helper reads from the same in-memory
          // mock cache.
          [kpiData, userData, riskData, trend, top] = await Promise.all([
            ReportService.getDashboardKpis(),
            OneDriveService.getAllOneDriveUsers(),
            ReportService.getGovernanceRisks(),
            OneDriveService.getStorageTrend(),
            OneDriveService.getTopOneDrives(10)
          ]);
        }

        if (!isMounted) {
          return;
        }

        // eslint-disable-next-line no-console
        console.log('Users loaded:', userData.length);
        // eslint-disable-next-line no-console
        console.log('KPI data loaded:', kpiData);
        // eslint-disable-next-line no-console
        console.log('DashboardDataContext: mapping successful for all sections (kpis, users, risks, storageTrend, topOneDrives).');

        setKpis(kpiData);
        setUsers(userData);
        setRisks(riskData);
        setStorageTrend(trend);
        setTopOneDrives(top);
      } catch (e) {
        if (isMounted) {
          const message = e instanceof Error && e.message
            ? e.message
            : 'Unable to reach the OneDrive Dashboard data service. Please verify the Azure Function endpoint or contact your administrator.';
          // eslint-disable-next-line no-console
          console.error('DashboardDataContext: Dashboard API load FAILED.', e);
          setError(message);
        }
      } finally {
        if (isMounted) {
          setLoading(false);
        }
      }
    }

    load().catch(() => { /* handled above */ });

    return () => { isMounted = false; };
  }, [context, apiBaseUrl, apiResourceUri, useMockData, refreshToken]);

  // "Refresh Data" is the ONLY thing that invalidates the dashboard
  // cache/de-duplication state (see OneDriveService.invalidateCache) and
  // forces a new GET /api/onedrive-dashboard request. Tab switching or
  // re-rendering never calls invalidateCache.
  const refresh = React.useCallback(() => {
    OneDriveService.invalidateCache();
    setRefreshToken(t => t + 1);
  }, []);

  const value = React.useMemo<IDashboardDataState>(() => ({
    loading, error, kpis, users, risks, storageTrend, topOneDrives, refresh, tenantRootUrl
  }), [loading, error, kpis, users, risks, storageTrend, topOneDrives, refresh, tenantRootUrl]);

  return <DashboardDataContext.Provider value={value}>{children}</DashboardDataContext.Provider>;
};

export function useDashboardData(): IDashboardDataState {
  const ctx = React.useContext(DashboardDataContext);
  if (!ctx) {
    throw new Error('useDashboardData must be used within a DashboardDataProvider.');
  }
  return ctx;
}
