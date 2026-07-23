import * as React from 'react';
import { Pivot, PivotItem } from '@fluentui/react/lib/Pivot';
import { DefaultButton } from '@fluentui/react/lib/Button';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';

import { IOneDriveDashboardProps } from './IOneDriveDashboardProps';
import { ThemeProvider } from '../../../contexts/ThemeContext';
import { DashboardDataProvider, useDashboardData } from '../../../contexts/DashboardDataContext';
import { KpiSection } from '../../../components/kpi/KpiSection';
import { StorageAnalyticsSection } from '../../../components/analytics/StorageAnalyticsSection';
import { InventoryTable } from '../../../components/inventory/InventoryTable';
import { InactiveOneDrivesSection } from '../../../components/inactive/InactiveOneDrivesSection';
import { GovernanceSection } from '../../../components/governance/GovernanceSection';
import { ErrorState } from '../../../components/common';
import { ReportService } from '../../../services/ReportService';

import styles from './OneDriveDashboard.module.scss';

const DashboardShell: React.FC<{ description: string }> = ({ description }) => {
  const { loading, error, kpis, users, risks, storageTrend, topOneDrives, refresh } = useDashboardData();
  const [buckets, setBuckets] = React.useState<{ d30: typeof users; d60: typeof users; d90: typeof users } | undefined>(undefined);

  React.useEffect(() => {
    if (!loading) {
      ReportService.getInactiveBuckets().then(setBuckets).catch(() => setBuckets(undefined));
    }
  }, [loading]);

  return (
    <div>
      <div className={styles.header}>
        <div>
          <h2 className={styles.headerTitle}>OneDrive Dashboard</h2>
          <div className={styles.headerSubtitle}>{description || 'Monitor and manage OneDrive usage across your Microsoft 365 tenant.'}</div>
        </div>
        <DefaultButton iconProps={{ iconName: 'Refresh' }} text="Refresh Data" onClick={refresh} disabled={loading} />
      </div>

      {error && <ErrorState message={error} onRetry={refresh} />}

      {!error && (
        <Pivot className={styles.pivot}>
          <PivotItem headerText="Overview" itemIcon="ViewDashboard">
            <div className={styles.sectionWrapper}>
              <KpiSection kpis={kpis} loading={loading} />
            </div>
            <div className={styles.sectionWrapper}>
              <StorageAnalyticsSection users={users} topOneDrives={topOneDrives} storageTrend={storageTrend} loading={loading} />
            </div>
          </PivotItem>

          <PivotItem headerText="OneDrive Inventory" itemIcon="Table">
            <div className={styles.sectionWrapper}>
              {loading ? <Spinner size={SpinnerSize.large} label="Loading OneDrive inventory..." /> : <InventoryTable />}
            </div>
          </PivotItem>

          <PivotItem headerText="Inactive OneDrives" itemIcon="UserRemove">
            <div className={styles.sectionWrapper}>
              {loading ? <Spinner size={SpinnerSize.large} label="Loading inactivity data..." /> : <InactiveOneDrivesSection buckets={buckets} loading={loading} />}
            </div>
          </PivotItem>

          <PivotItem headerText="Security & Governance" itemIcon="Shield">
            <div className={styles.sectionWrapper}>
              {loading ? <Spinner size={SpinnerSize.large} label="Evaluating governance risks..." /> : <GovernanceSection risks={risks} loading={loading} />}
            </div>
          </PivotItem>
        </Pivot>
      )}
    </div>
  );
};

export default class OneDriveDashboard extends React.Component<IOneDriveDashboardProps> {
  public render(): React.ReactElement<IOneDriveDashboardProps> {
    const { hasTeamsContext, context, theme, isDarkTheme, useMockData, description, apiBaseUrl } = this.props;

    return (
      <section className={`${styles.oneDriveDashboard} ${hasTeamsContext ? styles.teams : ''}`}>
        <ThemeProvider theme={theme} isDarkTheme={isDarkTheme}>
          <DashboardDataProvider context={context} apiBaseUrl={apiBaseUrl} useMockData={useMockData}>
            <DashboardShell description={description} />
          </DashboardDataProvider>
        </ThemeProvider>
      </section>
    );
  }
}