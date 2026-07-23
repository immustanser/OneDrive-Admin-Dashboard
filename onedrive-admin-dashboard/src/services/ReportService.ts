import { OneDriveService } from './OneDriveService';
import { calculateGovernanceRisks } from '../utils/riskCalculator';
import { daysSince } from '../utils/formatters';
import {
  IDashboardKpis,
  IGovernanceRiskItem,
  IOneDriveUser
} from '../models';

/**
 * Aggregates OneDrive inventory data into dashboard-ready KPIs, sharing
 * analytics, and governance/security risk reports.
 *
 * When real data is in use, KPIs / sharing placeholders / governance
 * risks are computed server-side by the Azure Function (see
 * api/src/services/dashboardMapperService.ts) and simply passed through
 * here. In the local mock-data path (OneDriveService "useMockData"),
 * these are derived client-side from the generated sample users so the
 * workbench experience keeps working without the Function configured.
 */
export class ReportService {
  public static async getDashboardKpis(): Promise<IDashboardKpis> {
    const apiKpis = await OneDriveService.getKpiData();
    if (apiKpis) {
      return apiKpis;
    }

    const users = await OneDriveService.getAllOneDriveUsers();

    const totalStorageConsumed = users.reduce((sum, u) => sum + u.storageUsedGB, 0);
    const totalStorageAllocated = users.reduce((sum, u) => sum + u.storageQuotaGB, 0);
    const activeUsers = users.filter(u => u.isActive30Days).length;
    const inactive = users.filter(u => !u.isActive90Days).length;
    const sharedFiles = users.reduce((sum, u) => sum + u.sharedFilesCount, 0);
    const externalSharing = users.reduce((sum, u) => sum + u.externalSharedFilesCount, 0);

    return {
      totalOneDriveSites: users.length,
      totalUsersWithOneDrive: users.length,
      totalStorageConsumedGB: Math.round(totalStorageConsumed),
      totalStorageAllocatedGB: Math.round(totalStorageAllocated),
      activeUsersLast30Days: activeUsers,
      inactiveOneDrives: inactive,
      sharedFilesCount: sharedFiles,
      externalSharingCount: externalSharing
    };
  }

  public static async getGovernanceRisks(): Promise<IGovernanceRiskItem[]> {
    const apiRisks = await OneDriveService.getGovernanceRisksFromApi();
    if (apiRisks) {
      return apiRisks;
    }
    const users = await OneDriveService.getAllOneDriveUsers();
    return calculateGovernanceRisks(users);
  }

  public static async getInactiveBuckets(): Promise<{ d30: IOneDriveUser[]; d60: IOneDriveUser[]; d90: IOneDriveUser[] }> {
    const users = await OneDriveService.getAllOneDriveUsers();
    return {
      d30: users.filter(u => daysSince(u.lastActivityDate) > 30 && daysSince(u.lastActivityDate) <= 60),
      d60: users.filter(u => daysSince(u.lastActivityDate) > 60 && daysSince(u.lastActivityDate) <= 90),
      d90: users.filter(u => daysSince(u.lastActivityDate) > 90)
    };
  }
}
