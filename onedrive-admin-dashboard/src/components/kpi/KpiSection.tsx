import * as React from 'react';
import { KpiCard } from './KpiCard';
import { KpiCardSkeleton } from '../common';
import { IDashboardKpis, IKpiCardData } from '../../models';
import { formatGB, formatNumber, percentage } from '../../utils/formatters';
import styles from './KpiCard.module.scss';

export interface IKpiSectionProps {
  kpis: IDashboardKpis | undefined;
  loading: boolean;
}

function buildKpiCards(kpis: IDashboardKpis): IKpiCardData[] {
  const storagePct = percentage(kpis.totalStorageConsumedGB, kpis.totalStorageAllocatedGB);
  const activePct = percentage(kpis.activeUsersLast30Days, kpis.totalUsersWithOneDrive);

  return [
    { id: 'sites', title: 'Total OneDrive Sites', value: formatNumber(kpis.totalOneDriveSites), rawValue: kpis.totalOneDriveSites, icon: 'CloudUpload', status: 'info' },
    { id: 'users', title: 'Total Users with OneDrive', value: formatNumber(kpis.totalUsersWithOneDrive), rawValue: kpis.totalUsersWithOneDrive, icon: 'Contact', status: 'neutral' },
    { id: 'consumed', title: 'Total Storage Consumed', value: formatGB(kpis.totalStorageConsumedGB), rawValue: kpis.totalStorageConsumedGB, icon: 'Database', status: storagePct > 85 ? 'danger' : storagePct > 65 ? 'warning' : 'success', subtitle: `${storagePct}% of allocated` },
    { id: 'allocated', title: 'Total Storage Allocated', value: formatGB(kpis.totalStorageAllocatedGB), rawValue: kpis.totalStorageAllocatedGB, icon: 'Storage', status: 'info' },
    { id: 'active30', title: 'Active Users (Last 30 Days)', value: formatNumber(kpis.activeUsersLast30Days), rawValue: kpis.activeUsersLast30Days, icon: 'PeopleAdd', status: 'success', subtitle: `${activePct}% of users` },
    { id: 'inactive', title: 'Inactive OneDrives', value: formatNumber(kpis.inactiveOneDrives), rawValue: kpis.inactiveOneDrives, icon: 'UserRemove', status: kpis.inactiveOneDrives > 0 ? 'warning' : 'success' },
    { id: 'shared', title: 'Shared Files Count', value: formatNumber(kpis.sharedFilesCount), rawValue: kpis.sharedFilesCount, icon: 'Share', status: 'info' },
    { id: 'external', title: 'External Sharing Count', value: formatNumber(kpis.externalSharingCount), rawValue: kpis.externalSharingCount, icon: 'GlobeWarning', status: kpis.externalSharingCount > 50 ? 'danger' : 'warning' }
  ];
}

export const KpiSection: React.FC<IKpiSectionProps> = ({ kpis, loading }) => {
  if (loading || !kpis) {
    return (
      <div className={styles.grid}>
        {Array.from({ length: 8 }).map((_, i) => <KpiCardSkeleton key={i} />)}
      </div>
    );
  }

  const cards = buildKpiCards(kpis);

  return (
    <div className={styles.grid}>
      {cards.map(card => <KpiCard key={card.id} data={card} />)}
    </div>
  );
};
