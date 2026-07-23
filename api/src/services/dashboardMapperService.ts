import { parseCsv, getField } from '../utils/csvParser';
import { toNumber, bytesToGB, percentage, roundTo } from '../utils/numberUtils';
import { parseReportDate, daysSince, formatMonthKey, formatMonthLabel, nowIso } from '../utils/dateUtils';
import {
  OneDriveUsageAccountDetailCsvHeaders,
  IOneDriveUsageAccountDetailRow
} from '../models/IOneDriveUsageAccountDetail';
import {
  OneDriveUsageStorageCsvHeaders,
  IOneDriveUsageStorageRow
} from '../models/IOneDriveStorageTrend';
import {
  IOneDriveDashboardResponse,
  IOneDriveUser,
  IDashboardKpis,
  IGovernanceRiskItem,
  ITopOneDrive,
  IInactiveBuckets,
  IStorageTrendPoint,
  OneDriveStatus,
  RiskCategory
} from '../models/IOneDriveDashboardResponse';

const H = OneDriveUsageAccountDetailCsvHeaders;
const SH = OneDriveUsageStorageCsvHeaders;

export function parseAccountDetailCsv(csvText: string): IOneDriveUsageAccountDetailRow[] {
  const rows = parseCsv(csvText);
  return rows.map((row) => ({
    reportRefreshDate: getField(row, H.reportRefreshDate),
    siteUrl: getField(row, H.siteUrl),
    ownerDisplayName: getField(row, H.ownerDisplayName),
    ownerPrincipalName: getField(row, H.ownerPrincipalName),
    isDeleted: getField(row, H.isDeleted).trim().toLowerCase() === 'true',
    lastActivityDate: getField(row, H.lastActivityDate) || undefined,
    fileCount: toNumber(getField(row, H.fileCount)),
    activeFileCount: toNumber(getField(row, H.activeFileCount)),
    storageUsedBytes: toNumber(getField(row, H.storageUsedBytes)),
    storageAllocatedBytes: toNumber(getField(row, H.storageAllocatedBytes)),
    reportPeriod: getField(row, H.reportPeriod)
  }));
}

export function parseStorageCsv(csvText: string): IOneDriveUsageStorageRow[] {
  const rows = parseCsv(csvText);
  return rows.map((row) => ({
    reportRefreshDate: getField(row, SH.reportRefreshDate),
    siteType: getField(row, SH.siteType) || undefined,
    storageUsedBytes: toNumber(getField(row, SH.storageUsedBytes)),
    reportDate: getField(row, SH.reportDate) || getField(row, SH.reportRefreshDate),
    reportPeriod: getField(row, SH.reportPeriod)
  }));
}

function deriveStatus(lastActivity: Date | undefined): { status: OneDriveStatus; days: number } {
  const days = daysSince(lastActivity);
  if (days <= 30) {
    return { status: 'Active', days };
  }
  if (days <= 90) {
    return { status: 'Warning', days };
  }
  return { status: 'Inactive', days };
}

/**
 * Maps raw account detail CSV rows to the dashboard's IOneDriveUser shape.
 * Deleted OneDrive sites are excluded from the active inventory.
 *
 * IMPORTANT: this mapper only uses the bulk OneDrive usage report and
 * NEVER calls Microsoft Graph per-user (e.g. /users/{upn}) here. On large
 * tenants (thousands of OneDrive owners), enriching every inventory row
 * with a per-user profile/manager call causes Microsoft Graph to respond
 * with 429 Too Many Requests. department, jobTitle and manager are
 * therefore left as empty placeholders in this bulk path; the SPFx client
 * fetches those fields on demand, only for the rows currently visible on
 * screen, via GET /api/user-profile?upn=... (see functions/getUserProfile.ts
 * and services/graphUserProfileService.ts).
 * Sharing counts are also not present in this Graph report and are left
 * as genuine zero placeholders rather than fabricated values.
 */
export function mapToInventoryUsers(rows: IOneDriveUsageAccountDetailRow[]): IOneDriveUser[] {
  return rows
    .filter((r) => !r.isDeleted)
    .map((r) => {
      const lastActivity = parseReportDate(r.lastActivityDate);
      const { status, days } = deriveStatus(lastActivity);
      const storageUsedGB = bytesToGB(r.storageUsedBytes);
      const storageQuotaGB = bytesToGB(r.storageAllocatedBytes);

      const user: IOneDriveUser = {
        id: r.ownerPrincipalName || r.siteUrl,
        displayName: r.ownerDisplayName,
        email: r.ownerPrincipalName,
        department: '',
        jobTitle: '',
        oneDriveUrl: (r.siteUrl || '').trim(),
        storageUsedGB,
        storageQuotaGB,
        filesCount: r.fileCount,
        lastActivityDate: r.lastActivityDate || '',
        manager: '',
        status,
        sharedFilesCount: 0,
        externalSharedFilesCount: 0,
        anonymousLinksCount: 0,
        companyLinksCount: 0,
        isActive30Days: days <= 30,
        isActive60Days: days <= 60,
        isActive90Days: days <= 90
      };
      return user;
    });
}

export function computeKpis(users: IOneDriveUser[]): IDashboardKpis {
  const totalStorageConsumedGB = roundTo(users.reduce((sum, u) => sum + u.storageUsedGB, 0), 2);
  const totalStorageAllocatedGB = roundTo(users.reduce((sum, u) => sum + u.storageQuotaGB, 0), 2);

  return {
    totalOneDriveSites: users.length,
    totalUsersWithOneDrive: users.length,
    totalStorageConsumedGB,
    totalStorageAllocatedGB,
    activeUsersLast30Days: users.filter((u) => u.isActive30Days).length,
    inactiveOneDrives: users.filter((u) => u.status === 'Inactive').length,
    // Sharing metrics are not provided by the OneDrive usage reports used
    // here; kept as explicit placeholders (0) rather than fabricated.
    sharedFilesCount: 0,
    externalSharingCount: 0
  };
}

export function computeInactiveBuckets(users: IOneDriveUser[]): IInactiveBuckets {
  return {
    inactive30Days: users.filter((u) => !u.isActive30Days).length,
    inactive60Days: users.filter((u) => !u.isActive60Days).length,
    inactive90Days: users.filter((u) => !u.isActive90Days).length
  };
}

export function computeTopOneDrives(users: IOneDriveUser[], count: number = 10): ITopOneDrive[] {
  return [...users]
    .sort((a, b) => b.storageUsedGB - a.storageUsedGB)
    .slice(0, count)
    .map((u) => ({ displayName: u.displayName, storageUsedGB: u.storageUsedGB }));
}

/**
 * One governance risk per user, most severe rule wins, based purely on
 * activity/storage signals available from the real Graph report data
 * (no fabricated sharing/manager-based risks in the real-data path).
 */
export function computeGovernanceRisks(users: IOneDriveUser[]): IGovernanceRiskItem[] {
  const risks: IGovernanceRiskItem[] = [];

  users.forEach((u) => {
    const storagePercent = percentage(u.storageUsedGB, u.storageQuotaGB);
    const days = !u.isActive90Days
      ? 91
      : !u.isActive60Days
        ? 61
        : !u.isActive30Days
          ? 31
          : 0;

    let category: RiskCategory | undefined;
    let riskLevel: IGovernanceRiskItem['riskLevel'] | undefined;
    let details = '';

    if (!u.isActive90Days && storagePercent >= 80) {
      category = 'Inactive OneDrive';
      riskLevel = 'Critical';
      details = `Inactive for ${daysSince(parseReportDate(u.lastActivityDate))} days while consuming ${storagePercent}% of allocated storage.`;
    } else if (storagePercent >= 90) {
      category = 'High Storage Consumer';
      riskLevel = 'High';
      details = `Storage usage at ${storagePercent}% of allocated quota (${u.storageUsedGB} GB of ${u.storageQuotaGB} GB).`;
    } else if (!u.isActive60Days) {
      category = 'Inactive OneDrive';
      riskLevel = 'Medium';
      details = `No activity detected for ${daysSince(parseReportDate(u.lastActivityDate))} days.`;
    } else if (!u.isActive30Days) {
      category = 'Inactive OneDrive';
      riskLevel = 'Low';
      details = `No activity detected for ${daysSince(parseReportDate(u.lastActivityDate))} days.`;
    }

    if (category && riskLevel) {
      risks.push({
        id: u.id,
        userDisplayName: u.displayName,
        email: u.email,
        category,
        riskLevel,
        details,
        detectedDate: nowIso()
      });
    }
    void days;
  });

  return risks;
}

/**
 * Aggregates the ~180 daily rows from getOneDriveUsageStorage into one
 * point per calendar month (latest value observed within that month),
 * preserving the existing "Storage Trend (12 Months)" chart semantics.
 */
export function computeStorageTrend(rows: IOneDriveUsageStorageRow[]): IStorageTrendPoint[] {
  const latestByMonth = new Map<string, { date: Date; storageUsedGB: number; label: string }>();

  rows.forEach((r) => {
    const date = parseReportDate(r.reportDate);
    if (!date) {
      return;
    }
    const key = formatMonthKey(date);
    const storageUsedGB = bytesToGB(r.storageUsedBytes);
    const existing = latestByMonth.get(key);
    if (!existing || date > existing.date) {
      latestByMonth.set(key, { date, storageUsedGB, label: formatMonthLabel(date) });
    }
  });

  return Array.from(latestByMonth.values())
    .sort((a, b) => a.date.getTime() - b.date.getTime())
    .map((v) => ({ month: v.label, storageUsedGB: v.storageUsedGB }));
}

/**
 * Groups users' storage usage by department. Since the bulk usage report
 * does not include department (see mapToInventoryUsers), this will bucket
 * everyone under "Unknown" unless/until the client has separately enriched
 * users[] with per-user profile data — this function never triggers Graph
 * calls itself, it only aggregates whatever department values are already
 * present on the given users.
 */
export function computeStorageByDepartment(
  users: IOneDriveUser[]
): { department: string; storageUsedGB: number }[] {
  const map = new Map<string, number>();
  users.forEach((u) => {
    const department = u.department && u.department.trim() ? u.department.trim() : 'Unknown';
    map.set(department, (map.get(department) ?? 0) + u.storageUsedGB);
  });
  return Array.from(map.entries())
    .map(([department, storageUsedGB]) => ({ department, storageUsedGB: roundTo(storageUsedGB, 2) }))
    .sort((a, b) => b.storageUsedGB - a.storageUsedGB);
}

export function buildDashboardResponse(
  accountDetailCsv: string,
  storageCsv: string
): IOneDriveDashboardResponse {
  const accountRows = parseAccountDetailCsv(accountDetailCsv);
  const storageRows = parseStorageCsv(storageCsv);

  const users = mapToInventoryUsers(accountRows);

  return {
    kpiData: computeKpis(users),
    inventoryUsers: users,
    storageAnalytics: {
      storageTrend: computeStorageTrend(storageRows),
      topOneDrives: computeTopOneDrives(users, 10),
      storageByDepartment: computeStorageByDepartment(users)
    },
    sharingReportPlaceholders: {
      totalSharedFiles: 0,
      externalSharedFiles: 0,
      anonymousLinks: 0,
      companyLinks: 0,
      mostSharedUsers: []
    },
    inactiveOneDrives: computeInactiveBuckets(users),
    governanceRisks: computeGovernanceRisks(users),
    generatedDateTime: nowIso(),
    dataSource: 'Microsoft Graph Reports API'
  };
}
