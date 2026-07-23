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
