import { WebPartContext } from '@microsoft/sp-webpart-base';
import { GraphService, IOneDriveDashboardApiResponse, ISendReminderRequest, ISendReminderResponse } from './GraphService';
import { CacheManager } from '../utils/cacheManager';
import { generateMockOneDriveUsers, generateMockStorageTrend } from '../utils/mockDataGenerator';
import { IOneDriveUser, IPagedQuery, IPagedResult, IStorageTrendPoint, ITopOneDrive, IDashboardKpis, IGovernanceRiskItem } from '../models';

const CACHE_TTL = 10 * 60 * 1000;

/**
 * Data access layer for OneDrive inventory data.
 *
 * Real data comes from the secure Azure Function backend (see /api).
 * GET /api/onedrive-dashboard is fetched AT MOST ONCE per CACHE_TTL
 * window: getDashboardData() below is the single choke point for that
 * call, with:
 *   - an in-memory 10-minute cache (dashboardCache), and
 *   - in-flight request de-duplication (dashboardPromise) so that any
 *     number of concurrent callers (Overview, Inventory, Sharing,
 *     Governance, Storage Analytics, etc.) during app startup share the
 *     SAME outstanding network request instead of each firing their own.
 *
 * DashboardDataContext is the single place that actually triggers this
 * fetch (once, on mount); all tabs/components consume the result via
 * context rather than calling the API independently. Only the "Refresh
 * Data" button (DashboardDataContext.refresh -> invalidateCache) clears
 * the cache to force a new request.
 *
 * "useMockData" is an explicit, opt-in LOCAL DEVELOPMENT toggle only
 * (set via the web part property pane). It is never used as a silent
 * fallback in production - if the Azure Function is unreachable, the
 * error is propagated so the UI can show a friendly error state.
 */
export class OneDriveService {
  private static useMockData: boolean = false;
  private static dashboardCache?: { data: IOneDriveDashboardApiResponse; expiresAt: number };
  private static dashboardPromise?: Promise<IOneDriveDashboardApiResponse>;
  private static forceRefreshNext: boolean = false;

  public static init(context: WebPartContext, apiBaseUrl: string, apiResourceUri: string, useMockData: boolean = false): void {
    this.useMockData = useMockData;
    GraphService.init(context, apiBaseUrl, apiResourceUri);
    if (useMockData) {
      // eslint-disable-next-line no-console
      console.warn('OneDriveService: useMockData=true - using local sample data, NOT the live Azure Function API. This should only be enabled for local development.');
    } else {
      // eslint-disable-next-line no-console
      console.log('OneDriveService: initialized in LIVE mode - data will be fetched from the Azure Function API.');
    }
  }

  /**
   * Single choke point for GET /api/onedrive-dashboard. See class-level
   * comment for the caching + de-duplication strategy.
   */
  private static async getDashboardData(): Promise<IOneDriveDashboardApiResponse> {
    const forceRefresh = this.forceRefreshNext;
    this.forceRefreshNext = false;

    if (!forceRefresh && this.dashboardCache && this.dashboardCache.expiresAt > Date.now()) {
      // eslint-disable-next-line no-console
      console.log('Dashboard API response served from cache');
      return this.dashboardCache.data;
    }

    if (!forceRefresh && this.dashboardPromise) {
      // eslint-disable-next-line no-console
      console.log('Dashboard API request in progress, reusing existing promise');
      return this.dashboardPromise;
    }

    // eslint-disable-next-line no-console
    console.log('Dashboard API request started');

    this.dashboardPromise = GraphService.getDashboardData(forceRefresh)
      .then((data) => {
        this.dashboardCache = { data, expiresAt: Date.now() + CACHE_TTL };
        // eslint-disable-next-line no-console
        console.log('Dashboard API loaded successfully');
        return data;
      })
      .catch((err) => {
        // Allow a subsequent call to retry after a failure instead of
        // being stuck reusing a rejected promise forever.
        this.dashboardPromise = undefined;
        throw err;
      });

    return this.dashboardPromise;
  }

  /**
   * Public single-fetch entry point used by DashboardDataContext to load
   * the full dashboard payload exactly once and distribute it to all
   * tabs via context. Returns undefined in mock mode (mock data has no
   * single combined API response - individual mock methods below are
   * used instead).
   */
  public static async getDashboardSnapshot(): Promise<IOneDriveDashboardApiResponse | undefined> {
    if (this.useMockData) {
      return undefined;
    }
    return this.getDashboardData();
  }

  /**
   * True when the web part is configured to use local mock data instead
   * of the live Azure Function API. Used by consumers (e.g. the
   * Inventory grid) to skip per-user profile enrichment calls, since
   * mock users already have department/manager populated.
   */
  public static isUsingMockData(): boolean {
    return this.useMockData;
  }

  public static async getAllOneDriveUsers(): Promise<IOneDriveUser[]> {
    if (this.useMockData) {
      return CacheManager.getOrFetch<IOneDriveUser[]>(
        'onedrive_users_mock',
        async () => generateMockOneDriveUsers(),
        CACHE_TTL
      );
    }
    const data = await this.getDashboardData();
    return data.inventoryUsers;
  }

  public static async getPagedOneDriveUsers(query: IPagedQuery): Promise<IPagedResult<IOneDriveUser>> {
    let items = await this.getAllOneDriveUsers();

    if (query.searchText) {
      const term = query.searchText.toLowerCase();
      items = items.filter(u =>
        u.displayName.toLowerCase().includes(term) ||
        u.email.toLowerCase().includes(term) ||
        u.department.toLowerCase().includes(term) ||
        u.manager.toLowerCase().includes(term)
      );
    }

    if (query.statusFilter && query.statusFilter.length > 0) {
      items = items.filter(u => query.statusFilter?.indexOf(u.status) !== -1);
    }

    if (query.departmentFilter && query.departmentFilter.length > 0) {
      items = items.filter(u => query.departmentFilter?.indexOf(u.department) !== -1);
    }

    if (query.sortField === 'healthPercent') {
      // Health is a derived field (storageUsedGB / storageQuotaGB), not
      // a real property on IOneDriveUser, so it needs its own comparator
      // rather than the generic keyof-based one below.
      items = [...items].sort((a, b) => {
        const av = a.storageQuotaGB > 0 ? a.storageUsedGB / a.storageQuotaGB : 0;
        const bv = b.storageQuotaGB > 0 ? b.storageUsedGB / b.storageQuotaGB : 0;
        return query.sortDescending ? bv - av : av - bv;
      });
    } else if (query.sortField) {
      const field = query.sortField as keyof IOneDriveUser;
      items = [...items].sort((a, b) => {
        const av = a[field];
        const bv = b[field];
        if (typeof av === 'number' && typeof bv === 'number') {
          return query.sortDescending ? bv - av : av - bv;
        }
        const cmp = String(av).localeCompare(String(bv));
        return query.sortDescending ? -cmp : cmp;
      });
    }

    const totalCount = items.length;
    const start = (query.page - 1) * query.pageSize;
    const pageItems = items.slice(start, start + query.pageSize);

    return { items: pageItems, totalCount, page: query.page, pageSize: query.pageSize };
  }

  /**
   * Distinct, alphabetically sorted list of non-empty department values
   * across the full inventory dataset (not just the current page), for
   * populating the Inventory tab's Department filter dropdown.
   */
  public static async getDepartmentOptions(): Promise<string[]> {
    const users = await this.getAllOneDriveUsers();
    const departments = new Set<string>();
    users.forEach(u => {
      if (u.department && u.department.trim()) {
        departments.add(u.department.trim());
      }
    });
    return Array.from(departments).sort((a, b) => a.localeCompare(b));
  }

  public static async getTopOneDrives(count: number = 10): Promise<ITopOneDrive[]> {
    if (this.useMockData) {
      const users = await this.getAllOneDriveUsers();
      return [...users]
        .sort((a, b) => b.storageUsedGB - a.storageUsedGB)
        .slice(0, count)
        .map(u => ({ displayName: u.displayName, storageUsedGB: u.storageUsedGB }));
    }
    const data = await this.getDashboardData();
    return data.storageAnalytics.topOneDrives.slice(0, count);
  }

  /**
   * Sends an inactivity reminder email for a single inactive OneDrive
   * owner, via the secure Azure Function backend (see GraphService).
   * In mock-data mode (local dev only, no live Function App), this is
   * simulated as an immediate success without any network call, so the
   * UI flow can still be exercised locally.
   */
  public static async sendInactiveOneDriveReminder(payload: ISendReminderRequest): Promise<ISendReminderResponse> {
    if (this.useMockData) {
      return { success: true, message: 'Reminder email sent successfully.' };
    }
    return GraphService.sendInactiveOneDriveReminder(payload);
  }

  public static async getStorageTrend(): Promise<IStorageTrendPoint[]> {
    if (this.useMockData) {
      return CacheManager.getOrFetch('onedrive_storage_trend_mock', async () => generateMockStorageTrend(), CACHE_TTL);
    }
    const data = await this.getDashboardData();
    return data.storageAnalytics.storageTrend;
  }

  /**
   * Server-computed KPIs from the Azure Function response. Undefined in the
   * mock-data path, where ReportService derives KPIs client-side instead.
   */
  public static async getKpiData(): Promise<IDashboardKpis | undefined> {
    if (this.useMockData) {
      return undefined;
    }
    const data = await this.getDashboardData();
    return data.kpiData;
  }

  public static async getGovernanceRisksFromApi(): Promise<IGovernanceRiskItem[] | undefined> {
    if (this.useMockData) {
      return undefined;
    }
    const data = await this.getDashboardData();
    return data.governanceRisks;
  }

  public static async getInactiveUsers(thresholdDays: 30 | 60 | 90): Promise<IOneDriveUser[]> {
    const users = await this.getAllOneDriveUsers();
    const field = thresholdDays === 30 ? 'isActive30Days' : thresholdDays === 60 ? 'isActive60Days' : 'isActive90Days';
    return users.filter(u => !u[field]);
  }

  /**
   * Invalidates all cached dashboard data (real dashboard cache + mock
   * caches). Called ONLY by DashboardDataContext.refresh() (the "Refresh
   * Data" button) - never automatically, and never by individual
   * tabs/components.
   */
  public static invalidateCache(): void {
    this.dashboardCache = undefined;
    this.dashboardPromise = undefined;
    this.forceRefreshNext = true;
    CacheManager.clear('onedrive_users_mock');
    CacheManager.clear('onedrive_storage_trend_mock');
  }
}
