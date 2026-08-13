# OneDrive Dashboard - Azure Function Backend API

Secure backend that fronts the Microsoft Graph Reports API for the SPFx
OneDrive Dashboard web part. The SPFx client never calls Microsoft Graph
directly and never sees a client secret - it only calls this Function.

```
SPFx web part  --(HTTPS GET)-->  Azure Function  --(client credentials)-->  Entra ID  --(app-only token)-->  Microsoft Graph Reports API
```

## Endpoints

| Route | Method | Description |
|---|---|---|
| `/api/onedrive-dashboard` | GET | Full dashboard payload (KPIs, inventory, storage analytics, inactive buckets, governance risks, sharing placeholders). |
| `/api/onedrive-account-details` | GET | OneDrive inventory only. |
| `/api/onedrive-storage-trend` | GET | Monthly storage trend only. |
| `/api/send-inactive-onedrive-reminder` | POST | Sends an inactivity reminder email to one inactive OneDrive owner (Inactive OneDrives tab). |

## Local development

0. Install the Azure Functions Core Tools globally (not a project
   dependency): `npm install -g azure-functions-core-tools@4 --unsafe-perm true`
1. `npm install`
2. Copy the placeholder secret in `local.settings.json` with your own
   test client secret value **locally only** - never commit this file
   (it is already in `.gitignore`).
3. `npm run build`
4. `npm start` (runs `func start`, serves on `http://localhost:7071/api/...`)

## Required application settings

Configure these as Azure Function **Application Settings** (or Key Vault
references) - never hardcode them in source:

| Setting | Value | Notes |
|---|---|---|
| `TENANT_ID` | `160f867a-f41f-432c-b770-4a07803e4f5b` | Not secret. |
| `CLIENT_ID` | `4860a82b-fb77-41e4-a838-94fa2ea7d080` | Not secret. |
| `CLIENT_SECRET` | *(secret)* | **Store only in Function App Settings or Key Vault.** Example Key Vault reference: `@Microsoft.KeyVault(SecretUri=https://<vault-name>.vault.azure.net/secrets/<secret-name>/)` |
| `GRAPH_SCOPE` | `https://graph.microsoft.com/.default` | Client-credentials scope. |
| `REMINDER_SENDER_UPN` | `SPAdmin01@stewart.com` | Not secret. Mailbox that "Send Reminder" emails are sent from (see `graphMailService.ts`). Must be a real, licensed mailbox in the tenant. |

## Entra ID App Registration

The app registration (`CLIENT_ID` above) must have these **application
(not delegated)** permissions granted with admin consent:

- `Reports.Read.All`
- `Directory.Read.All`
- `User.Read.All`
- `Sites.Read.All`
- `Mail.Send` - required for the "Send Reminder" feature on the Inactive
  OneDrives tab (`POST /api/send-inactive-onedrive-reminder`, implemented
  in `graphMailService.ts`). **This is a sensitive permission**: as an
  application (not delegated) permission, `Mail.Send` allows this app
  registration to send mail as **any** mailbox in the tenant, not just
  one. This backend limits the actual blast radius in code by only ever
  calling `sendMail` for the single mailbox configured via the
  `REMINDER_SENDER_UPN` application setting (intended value:
  `SPAdmin01@stewart.com`, the SharePoint administration service
  account) - no code path accepts a caller-supplied "from" address.
  Admin consent is required in Entra ID before this permission takes
  effect; without it, `/api/send-inactive-onedrive-reminder` will fail
  with a Graph authorization error.

## Deployment

1. `npm run build`
2. Deploy the Function App (e.g. `func azure functionapp publish <app-name>`,
   VS Code Azure Functions extension, or CI/CD pipeline).
3. Set the application settings listed above in the Azure Portal
   (Function App > Configuration).
4. Configure CORS (Function App > CORS) to allow the SharePoint tenant
   origin(s) that will host the SPFx web part (e.g.
   `https://<tenant>.sharepoint.com`).
5. For production, restrict the function's `authLevel` (currently
   `anonymous` for ease of initial setup/testing) via a function key,
   Azure AD auth (Easy Auth), or API Management, and lock CORS down to
   only the required origins.
6. Copy the deployed Function App's base URL (e.g.
   `https://<app-name>.azurewebsites.net`) into the SPFx web part's
   **Azure Function API Base URL** property pane setting.

## Data source and mapping notes

- `getOneDriveUsageAccountDetail(period='D180')` -> OneDrive inventory,
  KPIs, top 10 largest OneDrives, inactive buckets, governance risks.
- `getOneDriveUsageStorage(period='D180')` -> monthly storage trend chart.
- Both endpoints respond with an HTTP 302 whose `Location` header points
  to a short-lived CSV download URL; `graphReportsService.ts` follows
  that redirect manually and parses the CSV.
- Fields not present in these two reports (department, job title,
  manager, sharing/link counts) are returned as genuine empty/zero
  placeholders rather than fabricated values. See
  `dashboardMapperService.ts` for the full mapping and governance risk
  rules.
