/**
 * Represents a single row returned by the Microsoft Graph
 * `getOneDriveUsageStorage(period='D180')` report (CSV format).
 * Header names are matched case-insensitively/tolerantly by the CSV parser
 * since exact wording was not fully confirmed for this report.
 */
export interface IOneDriveUsageStorageRow {
  reportRefreshDate: string;
  siteType: string | undefined;
  storageUsedBytes: number;
  reportDate: string;
  reportPeriod: string;
}

/** Candidate raw CSV header names as returned by the Graph Reports API. */
export const OneDriveUsageStorageCsvHeaders = {
  reportRefreshDate: 'Report Refresh Date',
  siteType: 'Site Type',
  storageUsedBytes: 'Storage Used (Byte)',
  reportDate: 'Report Date',
  reportPeriod: 'Report Period'
};

/** Aggregated point used for the dashboard's storage trend line chart. */
export interface IStorageTrendPoint {
  month: string;
  storageUsedGB: number;
}
