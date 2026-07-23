/**
 * Represents a single row returned by the Microsoft Graph
 * `getOneDriveUsageAccountDetail(period='D180')` report (CSV format).
 * Column names below match the exact CSV headers returned by Graph.
 */
export interface IOneDriveUsageAccountDetailRow {
  reportRefreshDate: string;
  siteUrl: string;
  ownerDisplayName: string;
  ownerPrincipalName: string;
  isDeleted: boolean;
  lastActivityDate: string | undefined;
  fileCount: number;
  activeFileCount: number;
  storageUsedBytes: number;
  storageAllocatedBytes: number;
  reportPeriod: string;
}

/** Raw CSV header names as returned by the Graph Reports API. */
export const OneDriveUsageAccountDetailCsvHeaders = {
  reportRefreshDate: 'Report Refresh Date',
  siteUrl: 'Site URL',
  ownerDisplayName: 'Owner Display Name',
  ownerPrincipalName: 'Owner Principal Name',
  isDeleted: 'Is Deleted',
  lastActivityDate: 'Last Activity Date',
  fileCount: 'File Count',
  activeFileCount: 'Active File Count',
  storageUsedBytes: 'Storage Used (Byte)',
  storageAllocatedBytes: 'Storage Allocated (Byte)',
  reportPeriod: 'Report Period'
};
