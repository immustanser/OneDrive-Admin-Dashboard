export function toNumber(value: string | undefined | null): number {
  if (value === undefined || value === null || value.trim() === '') {
    return 0;
  }
  const n = Number(value);
  return isNaN(n) ? 0 : n;
}

export function bytesToGB(bytes: number): number {
  return roundTo(bytes / 1024 / 1024 / 1024, 2);
}

export function bytesToTB(bytes: number): number {
  return roundTo(bytes / 1024 / 1024 / 1024 / 1024, 2);
}

export function percentage(part: number, total: number): number {
  if (!total || total <= 0) {
    return 0;
  }
  return roundTo((part / total) * 100, 2);
}

export function roundTo(value: number, decimals: number): number {
  const factor = Math.pow(10, decimals);
  return Math.round(value * factor) / factor;
}
