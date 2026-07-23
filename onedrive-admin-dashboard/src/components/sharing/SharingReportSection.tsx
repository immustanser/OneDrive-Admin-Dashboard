import * as React from 'react';
import { PieChartCard } from '../../charts';
import { SectionCard } from '../common';
import { ISharingReport } from '../../models';
import { formatNumber } from '../../utils/formatters';
import styles from './SharingReportSection.module.scss';

export interface ISharingReportSectionProps {
  report: ISharingReport | undefined;
  loading: boolean;
}

export const SharingReportSection: React.FC<ISharingReportSectionProps> = ({ report, loading }) => {
  if (loading || !report) {
    return null;
  }

  const tiles = [
    { label: 'Total Shared Files', value: report.totalSharedFiles },
    { label: 'External Shared Files', value: report.externalSharedFiles },
    { label: 'Anonymous Links', value: report.anonymousLinks },
    { label: 'Company Links', value: report.companyLinks }
  ];

  return (
    <div>
      <div className={styles.tilesGrid}>
        {tiles.map(tile => (
          <div className={styles.tile} key={tile.label}>
            <div className={styles.tileValue}>{formatNumber(tile.value)}</div>
            <div className={styles.tileLabel}>{tile.label}</div>
          </div>
        ))}
      </div>

      <div className={styles.layout}>
        <PieChartCard
          title="Sharing Composition"
          labels={['External', 'Anonymous', 'Company']}
          data={[report.externalSharedFiles, report.anonymousLinks, report.companyLinks]}
        />
        <SectionCard title="Most Shared Users (Drill-Down)" icon="People">
          <div className={styles.drillDownCard} style={{ boxShadow: 'none', padding: 0 }}>
            {report.mostSharedUsers.map(u => (
              <div className={styles.userRow} key={u.email}>
                <span>{u.displayName} <span style={{ color: '#8a8886' }}>({u.email})</span></span>
                <strong>{formatNumber(u.sharedFilesCount)} files</strong>
              </div>
            ))}
          </div>
        </SectionCard>
      </div>
    </div>
  );
};
