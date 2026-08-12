import * as React from 'react';
import { DetailsList, IColumn, IDetailsHeaderProps, DetailsListLayoutMode, ConstrainMode, SelectionMode } from '@fluentui/react/lib/DetailsList';
import { ScrollablePane, ScrollbarVisibility } from '@fluentui/react/lib/ScrollablePane';
import { Sticky, StickyPositionType } from '@fluentui/react/lib/Sticky';
import { IRenderFunction } from '@fluentui/react/lib/Utilities';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { DefaultButton } from '@fluentui/react/lib/Button';
import { TooltipHost } from '@fluentui/react/lib/Tooltip';
import { EmptyState, HealthBadge } from '../common';
import { IOneDriveUser } from '../../models';
import { daysSince, formatNumber, formatStorageAuto } from '../../utils/formatters';
import { getHealthLevel } from '../../utils/health';
import { OneDriveService, UserProfileService } from '../../services';
import commonStyles from '../common/Common.module.scss';
import styles from './InactiveOneDrivesSection.module.scss';

export interface IInactiveOneDrivesSectionProps {
  buckets: { d30: IOneDriveUser[]; d60: IOneDriveUser[]; d90: IOneDriveUser[] } | undefined;
  loading: boolean;
}

type BucketKey = 'd30' | 'd60' | 'd90';
type ProfileStatus = 'loading' | 'error' | 'done';

const TABS: { key: BucketKey; label: string }[] = [
  { key: 'd30', label: 'Inactive 30-60 Days' },
  { key: 'd60', label: 'Inactive 60-90 Days' },
  { key: 'd90', label: 'Inactive 90+ Days' }
];

const PAGE_SIZE_OPTIONS: IDropdownOption[] = [
  { key: 10, text: '10 / page' },
  { key: 20, text: '20 / page' },
  { key: 50, text: '50 / page' },
  { key: 100, text: '100 / page' }
];

function sumStorageGB(users: IOneDriveUser[]): number {
  return users.reduce((sum, u) => sum + u.storageUsedGB, 0);
}

export const InactiveOneDrivesSection: React.FC<IInactiveOneDrivesSectionProps> = ({ buckets, loading }) => {
  const [activeTab, setActiveTab] = React.useState<BucketKey>('d90');
  const [page, setPage] = React.useState<number>(1);
  const [pageSize, setPageSize] = React.useState<number>(10);
  // Default sort: Storage Used DESC, so the highest-value cleanup
  // opportunities are immediately visible.
  const [sortField, setSortField] = React.useState<string>('storageUsedGB');
  const [sortDescending, setSortDescending] = React.useState<boolean>(true);
  // Local override map of email -> manager name, populated on demand
  // (see effect below) only for rows currently visible on this page -
  // same lazy, cached, visible-rows-only pattern already used by the
  // Inventory tab (see hooks/usePagedOneDriveUsers.ts). Department does
  // NOT need this treatment: it is already populated tenant-wide in the
  // cached dashboard payload, so no per-row lookup is required for it.
  const [managerByEmail, setManagerByEmail] = React.useState<Map<string, string>>(new Map());
  const [profileStatus, setProfileStatus] = React.useState<Map<string, ProfileStatus>>(new Map());

  const setSort = React.useCallback((field: string) => {
    setSortField(prevField => {
      if (prevField === field) {
        setSortDescending(prev => !prev);
      } else {
        setSortDescending(false);
      }
      return field;
    });
  }, []);

  const activeUsers = buckets ? buckets[activeTab] : [];

  const sortedUsers = React.useMemo(() => {
    const getSortValue = (u: IOneDriveUser): number | string => {
      switch (sortField) {
        case 'healthPercent':
          return u.storageQuotaGB > 0 ? u.storageUsedGB / u.storageQuotaGB : 0;
        case 'daysInactive':
          return daysSince(u.lastActivityDate);
        case 'manager':
          return (managerByEmail.get(u.email) || u.manager || '').toLowerCase();
        default:
          return u[sortField as keyof IOneDriveUser] as number | string;
      }
    };

    return [...activeUsers].sort((a, b) => {
      const av = getSortValue(a);
      const bv = getSortValue(b);
      if (typeof av === 'number' && typeof bv === 'number') {
        return sortDescending ? bv - av : av - bv;
      }
      const cmp = String(av).localeCompare(String(bv));
      return sortDescending ? -cmp : cmp;
    });
  }, [activeUsers, sortField, sortDescending, managerByEmail]);

  const totalCount = sortedUsers.length;
  const totalPages = Math.max(1, Math.ceil(totalCount / pageSize));
  const pageItems = React.useMemo(
    () => sortedUsers.slice((page - 1) * pageSize, page * pageSize),
    [sortedUsers, page, pageSize]
  );
  const visibleEmailsKey = pageItems.map(u => u.email).join('|');

  React.useEffect(() => { setPage(1); }, [activeTab]);

  React.useEffect(() => {
    if (page > totalPages) {
      setPage(totalPages);
    }
  }, [page, totalPages]);

  React.useEffect(() => {
    if (OneDriveService.isUsingMockData()) {
      return;
    }
    let isMounted = true;
    const visibleEmails = pageItems.map(u => u.email).filter(Boolean);
    const missing = visibleEmails.filter(e => !managerByEmail.has(e));
    if (missing.length === 0) {
      return;
    }

    setProfileStatus(prev => {
      const next = new Map(prev);
      missing.forEach(e => next.set(e, 'loading'));
      return next;
    });

    UserProfileService.getUserProfiles(missing)
      .then(profiles => {
        if (!isMounted) {
          return;
        }
        setManagerByEmail(prev => {
          const next = new Map(prev);
          missing.forEach(e => next.set(e, profiles.get(e)?.manager || ''));
          return next;
        });
        setProfileStatus(prev => {
          const next = new Map(prev);
          missing.forEach(e => next.set(e, UserProfileService.isProfileError(e) ? 'error' : 'done'));
          return next;
        });
      })
      .catch(() => {
        if (!isMounted) {
          return;
        }
        setProfileStatus(prev => {
          const next = new Map(prev);
          missing.forEach(e => next.set(e, 'error'));
          return next;
        });
      });

    return () => { isMounted = false; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [visibleEmailsKey]);

  const bucketSummary = React.useMemo(() => {
    const d30 = { count: buckets?.d30.length ?? 0, storageGB: sumStorageGB(buckets?.d30 ?? []) };
    const d60 = { count: buckets?.d60.length ?? 0, storageGB: sumStorageGB(buckets?.d60 ?? []) };
    const d90 = { count: buckets?.d90.length ?? 0, storageGB: sumStorageGB(buckets?.d90 ?? []) };
    return { d30, d60, d90, totalStorageGB: d30.storageGB + d60.storageGB + d90.storageGB };
  }, [buckets]);

  const onRenderDetailsHeader: IRenderFunction<IDetailsHeaderProps> = (props, defaultRender) => {
    if (!props || !defaultRender) {
      return null;
    }
    return (
      <Sticky stickyPosition={StickyPositionType.Header} isScrollSynced>
        {defaultRender(props)}
      </Sticky>
    );
  };

  const columns: IColumn[] = React.useMemo(() => {
    const makeColumn = (key: keyof IOneDriveUser, name: string, minWidth: number, maxWidth: number, render?: (item: IOneDriveUser) => React.ReactNode): IColumn => ({
      key,
      name,
      fieldName: key,
      minWidth,
      maxWidth,
      isResizable: true,
      isSorted: sortField === key,
      isSortedDescending: sortDescending,
      onColumnClick: () => setSort(key),
      onRender: render ? (item: IOneDriveUser) => render(item) : undefined
    });

    const renderManager = (item: IOneDriveUser): React.ReactNode => {
      const status = profileStatus.get(item.email);
      if (status === 'loading') {
        return 'Loading...';
      }
      const manager = managerByEmail.get(item.email) || item.manager;
      return manager || 'N/A';
    };

    return [
      makeColumn('displayName', 'User Name', 140, 180),
      makeColumn('email', 'Email', 180, 240),
      {
        key: 'healthPercent',
        name: 'Health',
        minWidth: 100,
        maxWidth: 130,
        isResizable: true,
        isSorted: sortField === 'healthPercent',
        isSortedDescending: sortDescending,
        onColumnClick: () => setSort('healthPercent'),
        onRender: (item: IOneDriveUser) => <HealthBadge level={getHealthLevel(item.storageUsedGB, item.storageQuotaGB)} />
      },
      makeColumn('storageUsedGB', 'Storage Used', 110, 140, item => formatStorageAuto(item.storageUsedGB)),
      {
        key: 'daysInactive',
        name: 'Days Inactive',
        minWidth: 100,
        maxWidth: 130,
        isResizable: true,
        isSorted: sortField === 'daysInactive',
        isSortedDescending: sortDescending,
        onColumnClick: () => setSort('daysInactive'),
        onRender: (item: IOneDriveUser) => formatNumber(daysSince(item.lastActivityDate))
      },
      {
        key: 'manager',
        name: 'Manager',
        minWidth: 120,
        maxWidth: 160,
        isResizable: true,
        isSorted: sortField === 'manager',
        isSortedDescending: sortDescending,
        onColumnClick: () => setSort('manager'),
        onRender: (item: IOneDriveUser) => renderManager(item)
      },
      makeColumn('department', 'Department', 120, 160, item => item.department || 'N/A'),
      {
        key: 'action',
        name: 'Action',
        minWidth: 140,
        maxWidth: 160,
        isResizable: false,
        onRender: () => (
          <TooltipHost content="Coming Soon - Inactivity Email Notifications">
            <span>
              <DefaultButton text="Send Reminder" iconProps={{ iconName: 'Mail' }} disabled />
            </span>
          </TooltipHost>
        )
      }
    ];
  }, [sortField, sortDescending, setSort, profileStatus, managerByEmail]);

  if (loading || !buckets) {
    return null;
  }

  return (
    <div>
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

      <div className={styles.summaryGrid}>
        <div className={commonStyles.card}>
          <span className={styles.summaryLabel}>30-60 Days</span>
          <span className={styles.summaryValue}>{formatNumber(bucketSummary.d30.count)} OneDrives</span>
          <span className={styles.summarySub}>{formatStorageAuto(bucketSummary.d30.storageGB)} recoverable</span>
        </div>
        <div className={commonStyles.card}>
          <span className={styles.summaryLabel}>60-90 Days</span>
          <span className={styles.summaryValue}>{formatNumber(bucketSummary.d60.count)} OneDrives</span>
          <span className={styles.summarySub}>{formatStorageAuto(bucketSummary.d60.storageGB)} recoverable</span>
        </div>
        <div className={commonStyles.card}>
          <span className={styles.summaryLabel}>90+ Days</span>
          <span className={styles.summaryValue}>{formatNumber(bucketSummary.d90.count)} OneDrives</span>
          <span className={styles.summarySub}>{formatStorageAuto(bucketSummary.d90.storageGB)} recoverable</span>
        </div>
        <div className={commonStyles.card}>
          <span className={styles.summaryLabel}>Total Recoverable Storage</span>
          <span className={styles.summaryValue}>{formatStorageAuto(bucketSummary.totalStorageGB)}</span>
          <span className={styles.summarySub}>Across all inactivity buckets</span>
        </div>
      </div>

      {activeUsers.length === 0 ? (
        <EmptyState icon="Completed" title="No users in this bucket" description="Nothing to review here right now." />
      ) : (
        <>
          <div className={styles.gridScroll}>
            <ScrollablePane scrollbarVisibility={ScrollbarVisibility.auto}>
              <DetailsList
                items={pageItems}
                columns={columns}
                layoutMode={DetailsListLayoutMode.fixedColumns}
                constrainMode={ConstrainMode.unconstrained}
                selectionMode={SelectionMode.none}
                onRenderDetailsHeader={onRenderDetailsHeader}
                compact
              />
            </ScrollablePane>
          </div>

          <div className={styles.footer}>
            <span className={styles.footerInfo}>
              Showing {pageItems.length === 0 ? 0 : (page - 1) * pageSize + 1}-{Math.min(page * pageSize, totalCount)} of {formatNumber(totalCount)} inactive OneDrives
            </span>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
              <Dropdown
                options={PAGE_SIZE_OPTIONS}
                selectedKey={pageSize}
                onChange={(_, option) => option && setPageSize(option.key as number)}
                styles={{ dropdown: { width: 110 } }}
              />
              <DefaultButton iconProps={{ iconName: 'ChevronLeft' }} disabled={page <= 1} onClick={() => setPage(page - 1)} />
              <span style={{ fontSize: 13 }}>Page {page} of {totalPages}</span>
              <DefaultButton iconProps={{ iconName: 'ChevronRight' }} disabled={page >= totalPages} onClick={() => setPage(page + 1)} />
            </div>
          </div>
        </>
      )}
    </div>
  );
};
