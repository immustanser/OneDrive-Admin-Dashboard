/**
 * Graph Reports API returns dates as "YYYY-MM-DD" strings (sometimes empty
 * when there has been no activity at all).
 */
export function parseReportDate(value: string | undefined | null): Date | undefined {
  if (!value || value.trim().length === 0) {
    return undefined;
  }
  const d = new Date(value);
  return isNaN(d.getTime()) ? undefined : d;
}

export function daysSince(date: Date | undefined, referenceDate: Date = new Date()): number {
  if (!date) {
    return Number.POSITIVE_INFINITY;
  }
  const msPerDay = 1000 * 60 * 60 * 24;
  return Math.floor((referenceDate.getTime() - date.getTime()) / msPerDay);
}

export function formatMonthKey(date: Date): string {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  return `${year}-${month}`;
}

const MONTH_LABELS = [
  'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
  'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'
];

export function formatMonthLabel(date: Date): string {
  return `${MONTH_LABELS[date.getMonth()]} ${date.getFullYear()}`;
}

export function nowIso(): string {
  return new Date().toISOString();
}
