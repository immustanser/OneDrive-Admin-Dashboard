import { IOneDriveUser, IGovernanceRiskItem, RiskLevel } from '../models';
import { daysSince, percentage } from './formatters';

let counter = 0;
function nextId(): string {
  counter += 1;
  return `risk-${counter}`;
}

export function calculateGovernanceRisks(users: IOneDriveUser[]): IGovernanceRiskItem[] {
  const risks: IGovernanceRiskItem[] = [];
  const now = new Date().toISOString();

  users.forEach(user => {
    const quotaPct = percentage(user.storageUsedGB, user.storageQuotaGB);

    if (!user.manager) {
      risks.push({
        id: nextId(),
        userDisplayName: user.displayName,
        email: user.email,
        category: 'No Manager',
        riskLevel: 'Medium',
        details: 'No manager assigned in Azure AD profile.',
        detectedDate: now
      });
    }

    if (quotaPct >= 95) {
      risks.push({
        id: nextId(),
        userDisplayName: user.displayName,
        email: user.email,
        category: 'High Storage Consumer',
        riskLevel: 'High',
        details: `Using ${quotaPct}% of allocated storage quota.`,
        detectedDate: now
      });
    }

    if (user.externalSharedFilesCount > 25) {
      risks.push({
        id: nextId(),
        userDisplayName: user.displayName,
        email: user.email,
        category: 'Excessive Sharing',
        riskLevel: 'High',
        details: `${user.externalSharedFilesCount} externally shared files detected.`,
        detectedDate: now
      });
    }

    if (user.anonymousLinksCount > 5) {
      const level: RiskLevel = user.anonymousLinksCount > 10 ? 'Critical' : 'High';
      risks.push({
        id: nextId(),
        userDisplayName: user.displayName,
        email: user.email,
        category: 'External Sharing Violation',
        riskLevel: level,
        details: `${user.anonymousLinksCount} anonymous sharing links active.`,
        detectedDate: now
      });
    }

    if (quotaPct >= 85 && quotaPct < 95) {
      risks.push({
        id: nextId(),
        userDisplayName: user.displayName,
        email: user.email,
        category: 'Quota Nearing Limit',
        riskLevel: 'Medium',
        details: `Storage quota at ${quotaPct}%.`,
        detectedDate: now
      });
    }

    if (daysSince(user.lastActivityDate) > 90 && user.filesCount > 500) {
      risks.push({
        id: nextId(),
        userDisplayName: user.displayName,
        email: user.email,
        category: 'Retention Policy Mismatch',
        riskLevel: 'Low',
        details: 'Large inactive drive not covered by an active retention policy.',
        detectedDate: now
      });
    }
  });

  const order: Record<RiskLevel, number> = { Critical: 0, High: 1, Medium: 2, Low: 3 };
  return risks.sort((a, b) => order[a.riskLevel] - order[b.riskLevel]);
}
