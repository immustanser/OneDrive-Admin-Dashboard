# OneDrive Dashboard (SPFx Web Part)

## Summary

**OneDrive Dashboard** is an enterprise-grade SharePoint Framework (SPFx) web part that gives SharePoint/M365 administrators a centralized view of OneDrive for Business usage, storage analytics, sharing behavior, inactive accounts, and governance/security risk across the tenant.

Built with React, TypeScript, Fluent UI, PnPjs v4 (Graph), and Chart.js, it renders KPI cards, interactive charts, a sortable/filterable/paginated inventory grid, a sharing report, inactive-OneDrive triage tools, and a governance risk register — all themed to match SharePoint's light/dark theme and the tenant's Fluent design language.

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.21.1-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- SharePoint Online (Modern pages, Teams tab, Teams Personal App, SharePoint full page)

## Prerequisites

- Node.js `>=22.14.0 <23.0.0`
- SharePoint Online tenant with admin access (to grant API permissions and deploy to the App Catalog)
- Global/SharePoint admin rights to approve Microsoft Graph API permission requests

## Solution

| Solution                    | Author(s)                                |
| ---------------------------- | ----------------------------------------- |
| onedrive-admin-dashboard     | Senior M365 Solution Architect / SPFx Team |

## Version history

| Version | Date         | Comments        |
| ------- | ------------ | --------------- |
| 1.0     | July 2026    | Initial release |

---

## 1. Architecture Overview

```
                         ┌─────────────────────────────────────────┐
                         │        SharePoint Online Page /          │
                         │        Teams Tab / Teams Personal App    │
                         └────────────────────┬──────────────────────┘
                                              │
                                  ┌───────────▼────────────┐
                                  │  OneDriveDashboardWebPart│  (SPFx WebPart, property pane)
                                  └───────────┬────────────┘
                                              │ props: context, theme, useMockData
                                  ┌───────────▼────────────┐
                                  │    OneDriveDashboard     │  React root component
                                  │  (Pivot: 5 dashboard tabs)│
                                  └───────────┬────────────┘
                       ┌───────────────────────┼───────────────────────┐
                       │                       │                       │
             ┌─────────▼─────────┐  ┌──────────▼──────────┐ ┌──────────▼──────────┐
             │   ThemeContext     │  │ DashboardDataContext │ │  Presentation Layer │
             │ (light/dark theme) │  │ (loading/error/data) │ │ components/ charts/ │
             └────────────────────┘  └──────────┬───────────┘ └──────────────────────┘
                                                 │
                                     ┌───────────▼────────────┐
                                     │      Service Layer      │
                                     │  services/               │
                                     │   - GraphService.ts       │
                                     │   - OneDriveService.ts    │
                                     │   - ReportService.ts      │
                                     └───────────┬────────────┘
                                                 │
                            ┌────────────────────┼────────────────────┐
                            │                    │                    │
                  ┌─────────▼────────┐ ┌─────────▼─────────┐ ┌────────▼────────┐
                  │  PnPjs v4 Graph   │ │  MSGraphClientFactory│ │  CacheManager   │
                  │ (users, sites)    │ │ (reports endpoint)   │ │ (in-memory TTL) │
                  └─────────┬────────┘ └─────────┬─────────┘ └─────────────────┘
                            │                    │
                            └─────────┬──────────┘
                                       │
                              ┌────────▼─────────┐
                              │  Microsoft Graph  │
                              │       API         │
                              └───────────────────┘
```

**Data flow**: `DashboardDataProvider` initializes `OneDriveService` (which wires up `GraphService`) once on mount, fetches KPIs/users/sharing/risks/trend in parallel via `Promise.all`, and exposes them through the `useDashboardData()` hook to every section. Each service call is wrapped by `CacheManager` (5-minute TTL) to minimize redundant Graph calls; a **Refresh Data** button invalidates the cache on demand.

If Graph calls fail or return no usable report rows (e.g. missing permissions, local workbench testing), `OneDriveService` transparently falls back to deterministic, seeded **sample data** (`utils/mockDataGenerator.ts`) so the dashboard is always demonstrable. A property-pane toggle (**"Use sample data"**) lets admins force sample mode explicitly.

## 2. Folder Structure

```
src/
├─ components/
│  ├─ common/            # StatusBadge, RiskBadge, EmptyState, ErrorState, LoadingSkeleton, SectionCard
│  ├─ kpi/                # KpiCard, KpiSection (Section 1)
│  ├─ analytics/          # StorageAnalyticsSection (Section 2 - charts)
│  ├─ inventory/          # InventoryTable (Section 3 - data grid)
│  ├─ sharing/            # SharingReportSection (Section 4)
│  ├─ inactive/           # InactiveOneDrivesSection (Section 5)
│  └─ governance/         # GovernanceSection (Section 6)
├─ charts/                # Chart.js wrappers: Pie, Bar, Line, Doughnut + chartSetup.ts (registration, palette)
├─ services/
│  ├─ GraphService.ts     # Raw Graph access (PnPjs v4 + MSGraphClientFactory for report endpoints)
│  ├─ OneDriveService.ts  # OneDrive inventory data access, pagination/search/sort, caching, mock fallback
│  ├─ ReportService.ts    # Aggregates KPIs, sharing report, governance risks, inactivity buckets
│  └─ index.ts
├─ models/                # TypeScript interfaces (IOneDriveUser, IKpiData, ISharingReport, IGovernanceRisk, ...)
├─ hooks/                 # useDebounce, usePagedOneDriveUsers (client-side "server-style" pagination)
├─ contexts/              # ThemeContext (SPFx theme), DashboardDataContext (shared dashboard state)
├─ utils/                 # formatters, exportUtils (CSV/Excel), cacheManager, riskCalculator, mockDataGenerator
├─ styles/                # _variables.scss (shared color palette, spacing, shadows, radii)
└─ webparts/oneDriveDashboard/
   ├─ OneDriveDashboardWebPart.ts        # SPFx web part entry point + property pane
   └─ components/
      ├─ OneDriveDashboard.tsx           # React root: Pivot shell wiring all 6 sections
      └─ IOneDriveDashboardProps.ts
```

## 3. Dashboard Sections

1. **KPI Cards** — 8 color-coded, hoverable cards (Total Sites, Users, Storage Consumed/Allocated, Active Users 30d, Inactive Drives, Shared Files, External Sharing).
2. **Storage Analytics** — Pie (usage by department), horizontal Bar (Top 10 largest OneDrives), Line (12-month storage trend), Doughnut (Active vs Inactive).
3. **OneDrive Inventory** — Fluent UI `DetailsList` with search (debounced), status/column sorting, column resizing, sticky header (`ScrollablePane` + `Sticky`), client-side "server-style" pagination, CSV/Excel export.
4. **Sharing Report** — Summary tiles (Total/External/Anonymous/Company shared), composition pie chart, most-shared-users drill-down list.
5. **Inactive OneDrives** — Tabs for 30/60/90+ day inactivity buckets with Send Reminder, Generate Report (CSV), Open Profile actions.
6. **Security & Governance** — Auto-calculated risk register (no manager, high storage, excessive/anonymous sharing, quota nearing limit, retention mismatch) with Critical/High/Medium/Low badges and filtering.

## 4. Microsoft Graph Integration & Required Permissions

The web part requests the following **delegated/application** Graph permissions via `webApiPermissionRequests` in `config/package-solution.json`:

| Permission             | Purpose                                                     |
| ----------------------- | ------------------------------------------------------------ |
| `User.Read.All`         | Enumerate tenant users, job titles, departments               |
| `Sites.Read.All`        | Read OneDrive personal sites metadata                         |
| `Reports.Read.All`      | `getOneDriveUsageAccountDetail` storage/activity usage report |
| `Directory.Read.All`    | Resolve manager relationships (`/users/{id}/manager`)         |

### Granting permissions (SharePoint Admin Center)

1. Deploy the `.sppkg` to the tenant **App Catalog** (see Deployment below).
2. Go to **SharePoint Admin Center → Advanced → API access**.
3. Approve the four pending Graph API requests listed above.
4. Permissions typically propagate within a few minutes.

> **Note:** Until permissions are approved (or when running via `gulp serve` locally), the web part automatically renders using generated **sample data** so it remains fully demonstrable.

## 5. Deployment Instructions

### Build & Package

```
npm install
npm run build              # gulp bundle (debug) — compiles, lints, bundles
gulp bundle --ship
gulp package-solution --ship
```

This produces `sharepoint/solution/onedrive-admin-dashboard.sppkg`.

### Deploy to Tenant App Catalog

1. Upload `onedrive-admin-dashboard.sppkg` to your tenant/site App Catalog.
2. When prompted, click **Deploy** to trust the package.
3. Go to **SharePoint Admin Center → Advanced → API access** and approve the Graph permission requests (see above).
4. Add the **OneDrive Dashboard** app to a site (`Site Contents → New → App`).
5. Add the **OneDrive Dashboard** web part to a modern page.

### Local Development

```
npm install
gulp serve
```

Opens the SPFx workbench (`https://localhost:4321/temp/workbench.html`) with sample data (Graph calls are not available in local workbench).

## 6. Performance Considerations

- **Caching** — `CacheManager` caches Graph/report results in-memory for 5 minutes; explicit "Refresh Data" invalidates the cache.
- **Memoization** — `React.useMemo`/`useCallback` used for derived chart data, computed KPI cards, and grid columns to avoid unnecessary re-renders.
- **Debounced search** — inventory search input is debounced (300ms) before triggering a re-query.
- **Client-side "server-style" pagination** — `usePagedOneDriveUsers` mimics server pagination (page/pageSize/sort/filter) so the grid only renders the current page's rows regardless of tenant size.
- **Parallel data loading** — KPI, inventory, sharing, risk, and trend data are fetched concurrently via `Promise.all`.
- **Lazy/skeleton rendering** — loading skeletons for KPI cards and charts avoid layout shift while data loads.

## 7. Theming

The web part reads the SPFx `IReadonlyTheme` (light/dark/custom) via `onThemeChanged` and exposes it through `ThemeContext`. The custom SCSS palette (`styles/_variables.scss`) uses the requested brand colors (`#0A83AE` primary, `#0078D4` secondary, `#F8F9FA` background) while Fluent UI controls (buttons, dropdowns, DetailsList) automatically adapt to the SharePoint theme.

## 8. Extending to Live Graph Data

`OneDriveService.getAllOneDriveUsers()` currently calls `GraphService.getAllUsers()` and falls back to sample data. To wire up full production data:

1. Call `GraphService.getOneDriveUsageReport('D180')` to retrieve the CSV usage report (via `MSGraphClientFactory`, since `getOneDriveUsageAccountDetail` isn't exposed as a PnPjs v4 fluent extension).
2. Parse the CSV and join rows with `GraphService.getAllUsers()` (by UPN/email) and `GraphService.getManager()` for manager names.
3. Map the joined result into `IOneDriveUser[]` and replace the `generateMockOneDriveUsers()` fallback call in `OneDriveService`.

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**
