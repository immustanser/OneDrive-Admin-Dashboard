export interface IPagedQuery {
  page: number;
  pageSize: number;
  searchText?: string;
  sortField?: string;
  sortDescending?: boolean;
  statusFilter?: string[];
  departmentFilter?: string[];
}

export interface IPagedResult<T> {
  items: T[];
  totalCount: number;
  page: number;
  pageSize: number;
}
