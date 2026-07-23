export function formatGB(value: number): string {
  if (value >= 1024) {
    return `${(value / 1024).toFixed(2)} TB`;
  }
  return `${value.toFixed(1)} GB`;
}

export function formatNumber(value: number): string {
  return new Intl.NumberFormat('en-US').format(value);
}

export function formatDate(isoDate: string): string {
  if (!isoDate) {
    return 'N/A';
  }
  const d = new Date(isoDate);
  return d.toLocaleDateString('en-US', { year: 'numeric', month: 'short', day: 'numeric' });
}

export function daysSince(isoDate: string): number {
  if (!isoDate) {
    // No last activity date at all (e.g. a OneDrive that has never been
    // accessed) counts as maximally inactive rather than being silently
    // excluded from every inactivity bucket.
    return Number.POSITIVE_INFINITY;
  }
  const d = new Date(isoDate);
  if (isNaN(d.getTime())) {
    return Number.POSITIVE_INFINITY;
  }
  const diff = Date.now() - d.getTime();
  return Math.floor(diff / (1000 * 60 * 60 * 24));
}

export function percentage(part: number, total: number): number {
  if (total === 0) {
    return 0;
  }
  return Math.round((part / total) * 1000) / 10;
}
