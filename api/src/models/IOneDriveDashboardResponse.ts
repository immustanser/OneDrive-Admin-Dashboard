/**
 * These interfaces intentionally mirror the SPFx frontend models
 * (onedrive-admin-dashboard/src/models/*.ts) exactly so the client
 * can consume this response with little/no transformation.
 */

export type OneDriveStatus = 'Active' | 'Warning' | 'Inactive';

export interface IOneDriveUser {
  id: string;
  displayName: string;
  email: string;
  department: string;
  jobTitle: string;
  oneDriveUrl: string;
  storageUsedGB: number;
  storageQuotaGB: number;
  filesCount: number;
  lastActivityDate: string;
  manager: string;
  status: OneDriveStatus;
  sharedFilesCount: number;
  externalSharedFilesCount: number;
  anonymousLinksCount: number;
  companyLinksCount: number;
  isActive30Days: boolean;
  isActive60Days: boolean;
  isActive90Days: boolean;
}

export interface IDashboardKpis {
  totalOneDriveSites: number;
  totalUsersWithOneDrive: number;
  totalStorageConsumedGB: number;
  totalStorageAllocatedGB: number;
  activeUsersLast30Days: number;
  inactiveOneDrives: number;
  sharedFilesCount: number;
  externalSharingCount: number;
}

export interface IMostSharedUser {
  displayName: string;
  email: string;
  sharedFilesCount: number;
}

export interface ISharingReport {
  totalSharedFiles: number;
  externalSharedFiles: number;
  anonymousLinks: number;
  companyLinks: number;
  mostSharedUsers: IMostSharedUser[];
}

export type RiskLevel = 'Critical' | 'High' | 'Medium' | 'Low';

export type RiskCategory =
  | 'No Manager'
  | 'High Storage Consumer'
  | 'Excessive Sharing'
  | 'External Sharing Violation'
  | 'Quota Nearing Limit'
  | 'Retention Policy Mismatch'
  | 'Inactive OneDrive';

export interface IGovernanceRiskItem {
  id: string;
  userDisplayName: string;
  email: string;
  category: RiskCategory;
  riskLevel: RiskLevel;
  details: string;
  detectedDate: string;
}

export interface IStorageTrendPoint {
  month: string;
  storageUsedGB: number;
}

export interface ITopOneDrive {
  displayName: string;
  storageUsedGB: number;
}

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

/**
 * Full payload returned by GET /api/onedrive-dashboard
 */
export interface IOneDriveDashboardResponse {
  kpiData: IDashboardKpis;
  inventoryUsers: IOneDriveUser[];
  storageAnalytics: IStorageAnalytics;
  sharingReportPlaceholders: ISharingReport;
  inactiveOneDrives: IInactiveBuckets;
  governanceRisks: IGovernanceRiskItem[];
  generatedDateTime: string;
  dataSource: string;
}
