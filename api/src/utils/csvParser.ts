/**
 * Minimal, dependency-free, quote-aware CSV parser for Microsoft Graph
 * Reports API output. Returns an array of row objects keyed by the
 * (trimmed) header names found on the first line.
 */
export function parseCsv(csvText: string): Array<Record<string, string>> {
  const lines = splitLines(csvText);
  if (lines.length === 0) {
    return [];
  }

  const headers = parseCsvLine(lines[0]).map((h) => h.trim());
  const rows: Array<Record<string, string>> = [];

  for (let i = 1; i < lines.length; i++) {
    const line = lines[i];
    if (line.trim().length === 0) {
      continue;
    }
    const values = parseCsvLine(line);
    const row: Record<string, string> = {};
    for (let c = 0; c < headers.length; c++) {
      row[headers[c]] = (values[c] !== undefined ? values[c] : '').trim();
    }
    rows.push(row);
  }

  return rows;
}

function splitLines(text: string): string[] {
  return text.replace(/\r\n/g, '\n').split('\n');
}

function parseCsvLine(line: string): string[] {
  const values: string[] = [];
  let current = '';
  let inQuotes = false;

  for (let i = 0; i < line.length; i++) {
    const ch = line[i];

    if (inQuotes) {
      if (ch === '"') {
        if (line[i + 1] === '"') {
          current += '"';
          i++;
        } else {
          inQuotes = false;
        }
      } else {
        current += ch;
      }
    } else {
      if (ch === '"') {
        inQuotes = true;
      } else if (ch === ',') {
        values.push(current);
        current = '';
      } else {
        current += ch;
      }
    }
  }
  values.push(current);
  return values;
}

/**
 * Tolerant, case-insensitive header lookup helper. Graph report CSV
 * headers are generally stable, but this guards against minor casing
 * or whitespace variations between report versions.
 */
export function getField(row: Record<string, string>, headerName: string): string {
  if (row[headerName] !== undefined) {
    return row[headerName];
  }
  const normalizedTarget = headerName.trim().toLowerCase();
  const matchKey = Object.keys(row).find((k) => k.trim().toLowerCase() === normalizedTarget);
  return matchKey ? row[matchKey] : '';
}
