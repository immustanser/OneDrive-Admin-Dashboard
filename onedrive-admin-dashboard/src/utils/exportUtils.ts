import { IOneDriveUser } from '../models';

const COLUMNS: { key: keyof IOneDriveUser; label: string }[] = [
  { key: 'displayName', label: 'User Name' },
  { key: 'email', label: 'Email' },
  { key: 'department', label: 'Department' },
  { key: 'oneDriveUrl', label: 'OneDrive URL' },
  { key: 'storageUsedGB', label: 'Storage Used (GB)' },
  { key: 'storageQuotaGB', label: 'Storage Quota (GB)' },
  { key: 'filesCount', label: 'Files Count' },
  { key: 'lastActivityDate', label: 'Last Activity Date' },
  { key: 'manager', label: 'Manager' },
  { key: 'status', label: 'Status' }
];

function downloadBlob(content: string | Blob, fileName: string, mimeType: string): void {
  const blob = typeof content === 'string' ? new Blob([content], { type: mimeType }) : content;
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

function escapeCsvValue(value: unknown): string {
  const str = String(value ?? '');
  if (str.includes(',') || str.includes('"') || str.includes('\n')) {
    return `"${str.replace(/"/g, '""')}"`;
  }
  return str;
}

export function exportToCsv(data: IOneDriveUser[], fileName: string = 'onedrive-inventory.csv'): void {
  const header = COLUMNS.map(c => escapeCsvValue(c.label)).join(',');
  const rows = data.map(row => COLUMNS.map(c => escapeCsvValue(row[c.key])).join(','));
  const csv = [header, ...rows].join('\r\n');
  downloadBlob('\ufeff' + csv, fileName, 'text/csv;charset=utf-8;');
}

export function exportToExcel(data: IOneDriveUser[], fileName: string = 'onedrive-inventory.xls'): void {
  const header = COLUMNS.map(c => `<th>${c.label}</th>`).join('');
  const rows = data
    .map(row => `<tr>${COLUMNS.map(c => `<td>${row[c.key]}</td>`).join('')}</tr>`)
    .join('');
  const table = `
    <html xmlns:x="urn:schemas-microsoft-com:office:excel">
      <head>
        <meta charset="utf-8" />
        <!--[if gte mso 9]>
        <xml>
          <x:ExcelWorkbook>
            <x:ExcelWorksheets>
              <x:ExcelWorksheet>
                <x:Name>OneDrive Inventory</x:Name>
                <x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions>
              </x:ExcelWorksheet>
            </x:ExcelWorksheets>
          </x:ExcelWorkbook>
        </xml>
        <![endif]-->
      </head>
      <body>
        <table border="1">
          <thead><tr>${header}</tr></thead>
          <tbody>${rows}</tbody>
        </table>
      </body>
    </html>`;
  downloadBlob(table, fileName, 'application/vnd.ms-excel');
}
