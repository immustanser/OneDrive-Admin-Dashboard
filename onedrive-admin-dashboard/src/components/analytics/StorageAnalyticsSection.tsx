import * as React from 'react';
import { PieChartCard, BarChartCard, LineChartCard, DoughnutChartCard } from '../../charts';
import { ChartCardSkeleton } from '../common';
import { IOneDriveUser, IStorageTrendPoint, ITopOneDrive } from '../../models';
import { UserProfileService } from '../../services';
import chartStyles from '../../charts/ChartCard.module.scss';

export interface IStorageAnalyticsSectionProps {
  users: IOneDriveUser[];
  topOneDrives: ITopOneDrive[];
  storageTrend: IStorageTrendPoint[];
  loading: boolean;
}

/**
 * Groups users' storage usage by department using ONLY department values
 * already loaded (either directly on the user record, or previously
 * fetched/cached via the Inventory grid's per-visible-row profile lookups
 * - see UserProfileService). This never triggers new profile fetches for
 * the full tenant; users with no known department yet are bucketed under
 * "Unknown" rather than being enriched on demand here.
 */
function buildDepartmentDistribution(users: IOneDriveUser[]): { labels: string[]; data: number[] } {
  const map = new Map<string, number>();
  users.forEach(u => {
    const cachedProfile = UserProfileService.getCachedProfile(u.email);
    const rawDepartment = (cachedProfile && cachedProfile.department) || u.department;
    const department = rawDepartment && rawDepartment.trim() ? rawDepartment.trim() : 'Unknown';
    map.set(department, (map.get(department) ?? 0) + u.storageUsedGB);
  });
  const entries = Array.from(map.entries()).sort((a, b) => b[1] - a[1]);
  return { labels: entries.map(e => e[0]), data: entries.map(e => Math.round(e[1])) };
}

export const StorageAnalyticsSection: React.FC<IStorageAnalyticsSectionProps> = ({ users, topOneDrives, storageTrend, loading }) => {
  const distribution = React.useMemo(() => buildDepartmentDistribution(users), [users]);
  const activeCount = React.useMemo(() => users.filter(u => u.status === 'Active').length, [users]);
  const inactiveCount = users.length - activeCount;

  if (loading) {
    return (
      <div className={chartStyles.chartsGrid}>
        {Array.from({ length: 4 }).map((_, i) => <ChartCardSkeleton key={i} />)}
      </div>
    );
  }

  return (
    <div className={chartStyles.chartsGrid}>
      <PieChartCard title="Storage Usage Distribution (by Department)" labels={distribution.labels} data={distribution.data} />
      <BarChartCard
        title="Top 10 Largest OneDrives"
        labels={topOneDrives.map(t => t.displayName)}
        data={topOneDrives.map(t => Math.round(t.storageUsedGB))}
        horizontal
      />
      <LineChartCard title="Storage Trend (12 Months)" labels={storageTrend.map(t => t.month)} data={storageTrend.map(t => t.storageUsedGB)} />
      <DoughnutChartCard
        title="Active vs Inactive OneDrives"
        labels={['Active', 'Inactive/Warning']}
        data={[activeCount, inactiveCount]}
        colors={['#107C10', '#D13438']}
      />
    </div>
  );
};
