export type KpiStatus = 'success' | 'warning' | 'danger' | 'info' | 'neutral';

export interface IKpiCardData {
  id: string;
  title: string;
  value: string;
  rawValue: number;
  icon: string;
  status: KpiStatus;
  trend?: number;
  subtitle?: string;
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
