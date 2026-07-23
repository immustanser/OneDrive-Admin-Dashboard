import * as React from 'react';
import { DetailsList, IColumn, IDetailsHeaderProps, DetailsListLayoutMode, ConstrainMode, SelectionMode } from '@fluentui/react/lib/DetailsList';
import { ScrollablePane, ScrollbarVisibility } from '@fluentui/react/lib/ScrollablePane';
import { Sticky, StickyPositionType } from '@fluentui/react/lib/Sticky';
import { IRenderFunction } from '@fluentui/react/lib/Utilities';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { usePagedOneDriveUsers } from '../../hooks';
import { StatusBadge, EmptyState } from '../common';
import { formatGB, formatDate, formatNumber } from '../../utils/formatters';
import { exportToCsv, exportToExcel } from '../../utils/exportUtils';
import { IOneDriveUser } from '../../models';
import { useDashboardData } from '../../contexts/DashboardDataContext';
import { buildOneDriveFallbackUrl } from '../../utils/oneDriveUrl';
import styles from './InventoryTable.module.scss';

const STATUS_OPTIONS: IDropdownOption[] = [
  { key: 'Active', text: 'Active' },
  { key: 'Warning', text: 'Warning' },
  { key: 'Inactive', text: 'Inactive' }
];

const PAGE_SIZE_OPTIONS: IDropdownOption[] = [
  { key: 25, text: '25 / page' },
  { key: 50, text: '50 / page' },
  { key: 100, text: '100 / page' }
];

export const InventoryTable: React.FC = () => {
  const { tenantRootUrl } = useDashboardData();
  const {
    items, totalCount, loading, page, pageSize, setPage, setPageSize,
    searchText, setSearchText, sortField, sortDescending, setSort,
    statusFilter, setStatusFilter, profileStatus
  } = usePagedOneDriveUsers(25);

  const totalPages = Math.max(1, Math.ceil(totalCount / pageSize));

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

    const renderProfileField = (item: IOneDriveUser, value: string): React.ReactNode => {
      const status = profileStatus.get(item.email);
      if (status === 'loading') {
        return 'Loading...';
      }
      if (status === 'error') {
        return 'Unknown';
      }
      return value || '—';
    };

    return [
      makeColumn('displayName', 'User Name', 140, 180),
      makeColumn('email', 'Email', 180, 240),
      makeColumn('department', 'Department', 110, 150, item => renderProfileField(item, item.department)),
      makeColumn('oneDriveUrl', 'OneDrive URL', 160, 260, item => {
        const url = item.oneDriveUrl || buildOneDriveFallbackUrl(item.email, tenantRootUrl);
        return url
          ? <a href={url} target="_blank" rel="noreferrer" title={url}>Open Drive</a>
          : <span>N/A</span>;
      }),
      makeColumn('storageUsedGB', 'Storage Used (GB)', 110, 140, item => formatGB(item.storageUsedGB)),
      makeColumn('storageQuotaGB', 'Storage Quota (GB)', 110, 140, item => formatGB(item.storageQuotaGB)),
      makeColumn('filesCount', 'Files Count', 90, 120, item => formatNumber(item.filesCount)),
      makeColumn('lastActivityDate', 'Last Activity Date', 120, 150, item => formatDate(item.lastActivityDate)),
      makeColumn('manager', 'Manager', 120, 160, item => renderProfileField(item, item.manager)),
      makeColumn('status', 'Status', 100, 120, item => <StatusBadge status={item.status} />)
    ];
  }, [sortField, sortDescending, setSort, profileStatus, tenantRootUrl]);

  return (
    <div className={styles.container}>
      <div className={styles.toolbar}>
        <div className={styles.toolbarLeft}>
          <SearchBox
            className={styles.searchBox}
            placeholder="Search by name, email, department, manager..."
            value={searchText}
            onChange={(_, v) => setSearchText(v ?? '')}
          />
          <Dropdown
            placeholder="Filter status"
            multiSelect
            options={STATUS_OPTIONS}
            selectedKeys={statusFilter}
            onChange={(_, option) => {
              if (!option) { return; }
              const key = option.key as string;
              setStatusFilter(option.selected ? [...statusFilter, key] : statusFilter.filter(s => s !== key));
            }}
            styles={{ dropdown: { width: 160 } }}
          />
        </div>
        <div className={styles.toolbarRight}>
          <DefaultButton iconProps={{ iconName: 'ExcelDocument' }} text="Export Excel" onClick={() => exportToExcel(items, 'onedrive-inventory.xls')} />
          <PrimaryButton iconProps={{ iconName: 'Download' }} text="Export CSV" onClick={() => exportToCsv(items, 'onedrive-inventory.csv')} />
        </div>
      </div>

      <div className={styles.gridScroll}>
        {!loading && items.length === 0 ? (
          <EmptyState icon="SearchIssue" title="No matching OneDrives" description="Try adjusting your search or filters." />
        ) : (
          <ScrollablePane scrollbarVisibility={ScrollbarVisibility.auto}>
            <DetailsList
              items={items}
              columns={columns}
              layoutMode={DetailsListLayoutMode.fixedColumns}
              constrainMode={ConstrainMode.unconstrained}
              selectionMode={SelectionMode.none}
              onRenderDetailsHeader={onRenderDetailsHeader}
              compact
            />
          </ScrollablePane>
        )}
      </div>

      <div className={styles.footer}>
        <span className={styles.footerInfo}>
          Showing {items.length === 0 ? 0 : (page - 1) * pageSize + 1}-{Math.min(page * pageSize, totalCount)} of {formatNumber(totalCount)} OneDrives
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
    </div>
  );
};
