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
