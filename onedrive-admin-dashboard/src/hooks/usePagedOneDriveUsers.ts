import * as React from 'react';
import { OneDriveService, UserProfileService } from '../services';
import { IOneDriveUser, IPagedQuery } from '../models';
import { useDebounce } from './useDebounce';

export type ProfileStatus = 'loading' | 'error' | 'done';

export interface IUsePagedOneDriveUsersResult {
  items: IOneDriveUser[];
  totalCount: number;
  loading: boolean;
  page: number;
  pageSize: number;
  setPage: (page: number) => void;
  setPageSize: (size: number) => void;
  searchText: string;
  setSearchText: (text: string) => void;
  sortField: string | undefined;
  sortDescending: boolean;
  setSort: (field: string) => void;
  statusFilter: string[];
  setStatusFilter: (statuses: string[]) => void;
  /**
   * Per-user (keyed by email/UPN) status of the on-demand profile
   * enrichment call, for rows currently/previously visible on screen
   * only. Used by InventoryTable to render "Loading..." while pending
   * and "Unknown" if the lookup failed.
   */
  profileStatus: Map<string, ProfileStatus>;
}

/**
 * Encapsulates server-side style pagination, sorting, search and
 * status filtering for the OneDrive inventory grid. Backed by
 * OneDriveService.getPagedOneDriveUsers, which mimics a server call.
 */
export function usePagedOneDriveUsers(initialPageSize: number = 25): IUsePagedOneDriveUsersResult {
  const [items, setItems] = React.useState<IOneDriveUser[]>([]);
  const [totalCount, setTotalCount] = React.useState<number>(0);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [page, setPage] = React.useState<number>(1);
  const [pageSize, setPageSize] = React.useState<number>(initialPageSize);
  const [searchText, setSearchText] = React.useState<string>('');
  const [sortField, setSortField] = React.useState<string | undefined>('storageUsedGB');
  const [sortDescending, setSortDescending] = React.useState<boolean>(true);
  const [statusFilter, setStatusFilter] = React.useState<string[]>([]);
  const [profileStatus, setProfileStatus] = React.useState<Map<string, ProfileStatus>>(new Map());

  const debouncedSearch = useDebounce(searchText, 300);

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

  React.useEffect(() => {
    let isMounted = true;
    setLoading(true);

    const query: IPagedQuery = {
      page,
      pageSize,
      searchText: debouncedSearch,
      sortField,
      sortDescending,
      statusFilter
    };

    OneDriveService.getPagedOneDriveUsers(query)
      .then(result => {
        if (!isMounted) {
          return;
        }
        setItems(result.items);
        setTotalCount(result.totalCount);
        setLoading(false);

        // Enrich ONLY the currently visible page of rows with profile info
        // (department, manager, jobTitle) from GET /api/user-profile - e.g.
        // 25 calls for a page size of 25, never the full tenant. This is
        // what avoids Microsoft Graph 429 throttling on large tenants.
        if (!OneDriveService.isUsingMockData()) {
          const visibleUpns = result.items.map(u => u.email).filter(Boolean);

          setProfileStatus(prev => {
            const next = new Map(prev);
            visibleUpns.forEach(upn => next.set(upn, 'loading'));
            return next;
          });

          UserProfileService.getUserProfiles(visibleUpns)
            .then(profiles => {
              if (!isMounted) {
                return;
              }

              setProfileStatus(prev => {
                const next = new Map(prev);
                visibleUpns.forEach(upn => {
                  next.set(upn, UserProfileService.isProfileError(upn) ? 'error' : 'done');
                });
                return next;
              });

              if (profiles.size === 0) {
                return;
              }
              setItems(prevItems => prevItems.map(u => {
                const profile = profiles.get(u.email);
                if (!profile) {
                  return u;
                }
                return {
                  ...u,
                  department: profile.department || u.department,
                  jobTitle: profile.jobTitle || u.jobTitle,
                  manager: profile.manager || u.manager
                };
              }));
            })
            .catch(() => {
              if (!isMounted) {
                return;
              }
              setProfileStatus(prev => {
                const next = new Map(prev);
                visibleUpns.forEach(upn => next.set(upn, 'error'));
                return next;
              });
            });
        }
      })
      .catch(() => { if (isMounted) { setLoading(false); } });

    return () => { isMounted = false; };
  }, [page, pageSize, debouncedSearch, sortField, sortDescending, statusFilter]);

  React.useEffect(() => { setPage(1); }, [debouncedSearch, statusFilter]);

  return {
    items, totalCount, loading, page, pageSize, setPage, setPageSize,
    searchText, setSearchText, sortField, sortDescending, setSort,
    statusFilter, setStatusFilter, profileStatus
  };
}
