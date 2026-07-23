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

export interface ISharingDrillDownItem {
  fileName: string;
  owner: string;
  sharedWith: string;
  linkType: 'Anonymous' | 'Company' | 'Direct' | 'External';
  sharedDate: string;
}
