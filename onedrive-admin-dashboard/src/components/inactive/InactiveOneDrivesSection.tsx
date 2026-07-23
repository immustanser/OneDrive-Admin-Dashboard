import * as React from 'react';
import { DefaultButton } from '@fluentui/react/lib/Button';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { EmptyState } from '../common';
import { IOneDriveUser } from '../../models';
import { formatDate, daysSince } from '../../utils/formatters';
import { exportToCsv } from '../../utils/exportUtils';
import { useDashboardData } from '../../contexts/DashboardDataContext';
import styles from './InactiveOneDrivesSection.module.scss';

export interface IInactiveOneDrivesSectionProps {
  buckets: { d30: IOneDriveUser[]; d60: IOneDriveUser[]; d90: IOneDriveUser[] } | undefined;
  loading: boolean;
}

type BucketKey = 'd30' | 'd60' | 'd90';

const TABS: { key: BucketKey; label: string }[] = [
  { key: 'd30', label: 'Inactive 30-60 Days' },
  { key: 'd60', label: 'Inactive 60-90 Days' },
  { key: 'd90', label: 'Inactive 90+ Days' }
];

export const InactiveOneDrivesSection: React.FC<IInactiveOneDrivesSectionProps> = ({ buckets, loading }) => {
  const { tenantRootUrl } = useDashboardData();
  const [activeTab, setActiveTab] = React.useState<BucketKey>('d90');
  const [notice, setNotice] = React.useState<string | undefined>(undefined);

  if (loading || !buckets) {
    return null;
  }

  const activeUsers = buckets[activeTab];

  const sendReminder = (user: IOneDriveUser): void => setNotice(`Reminder email queued for ${user.displayName}.`);
  const generateReport = (user: IOneDriveUser): void => exportToCsv([user], `${user.displayName.replace(/\s+/g, '-')}-report.csv`);
  const openProfile = (user: IOneDriveUser): void => {
    const root = tenantRootUrl || 'https://contoso.sharepoint.com';
    window.open(`${root}/_layouts/15/userdisp.aspx?ID=${user.id}`, '_blank');
  };

  return (
    <div>
      {notice && (
        <MessageBar messageBarType={MessageBarType.success} onDismiss={() => setNotice(undefined)} isMultiline={false} styles={{ root: { marginBottom: 12 } }}>
          {notice}
        </MessageBar>
      )}

      <div className={styles.tabs}>
        {TABS.map(tab => (
          <div
            key={tab.key}
            className={`${styles.tab} ${activeTab === tab.key ? styles.tabActive : ''}`}
            onClick={() => setActiveTab(tab.key)}
          >
            {tab.label} ({buckets[tab.key].length})
          </div>
        ))}
      </div>

      {activeUsers.length === 0 ? (
        <EmptyState icon="Completed" title="No users in this bucket" description="Nothing to review here right now." />
      ) : (
        activeUsers.map(user => (
          <div className={styles.row} key={user.id}>
            <div className={styles.userInfo}>
              <span className={styles.userName}>{user.displayName}</span>
              <span className={styles.userMeta}>{user.email} • Last active {formatDate(user.lastActivityDate)} ({daysSince(user.lastActivityDate)} days ago)</span>
            </div>
            <div className={styles.actions}>
              <DefaultButton text="Send Reminder" iconProps={{ iconName: 'Mail' }} onClick={() => sendReminder(user)} />
              <DefaultButton text="Generate Report" iconProps={{ iconName: 'ReportDocument' }} onClick={() => generateReport(user)} />
              <DefaultButton text="Open Profile" iconProps={{ iconName: 'Contact' }} onClick={() => openProfile(user)} />
            </div>
          </div>
        ))
      )}
    </div>
  );
};
