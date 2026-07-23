import * as React from 'react';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { RiskBadge, EmptyState } from '../common';
import { IGovernanceRiskItem, RiskLevel } from '../../models';
import { formatDate } from '../../utils/formatters';
import styles from './GovernanceSection.module.scss';

export interface IGovernanceSectionProps {
  risks: IGovernanceRiskItem[];
  loading: boolean;
}

const RISK_OPTIONS: IDropdownOption[] = [
  { key: 'Critical', text: 'Critical' },
  { key: 'High', text: 'High' },
  { key: 'Medium', text: 'Medium' },
  { key: 'Low', text: 'Low' }
];

export const GovernanceSection: React.FC<IGovernanceSectionProps> = ({ risks, loading }) => {
  const [filter, setFilter] = React.useState<string[]>([]);

  if (loading) {
    return null;
  }

  const counts: Record<RiskLevel, number> = { Critical: 0, High: 0, Medium: 0, Low: 0 };
  risks.forEach(r => { counts[r.riskLevel] += 1; });

  const filtered = filter.length === 0 ? risks : risks.filter(r => filter.indexOf(r.riskLevel) !== -1);

  return (
    <div>
      <div className={styles.summaryGrid}>
        {(Object.keys(counts) as RiskLevel[]).map(level => (
          <div className={styles.summaryTile} key={level}>
            <div className={styles.summaryValue}>{counts[level]}</div>
            <div className={styles.summaryLabel}><RiskBadge level={level} /></div>
          </div>
        ))}
      </div>

      <Dropdown
        placeholder="Filter by risk level"
        multiSelect
        options={RISK_OPTIONS}
        selectedKeys={filter}
        onChange={(_, option) => {
          if (!option) { return; }
          const key = option.key as string;
          setFilter(option.selected ? [...filter, key] : filter.filter(s => s !== key));
        }}
        styles={{ dropdown: { width: 220, marginBottom: 12 } }}
      />

      {filtered.length === 0 ? (
        <EmptyState icon="Shield" title="No governance risks found" description="Your tenant looks compliant for the selected filters." />
      ) : (
        <div>
          <div className={`${styles.row} ${styles.headerRow}`}>
            <span>User</span>
            <span>Category</span>
            <span>Risk Level</span>
            <span>Details</span>
            <span>Detected</span>
          </div>
          {filtered.map(risk => (
            <div className={styles.row} key={risk.id}>
              <span>{risk.userDisplayName}</span>
              <span>{risk.category}</span>
              <span><RiskBadge level={risk.riskLevel} /></span>
              <span>{risk.details}</span>
              <span>{formatDate(risk.detectedDate)}</span>
            </div>
          ))}
        </div>
      )}
    </div>
  );
};
