export type HealthLevel = 'Healthy' | 'Warning' | 'High Risk' | 'Critical';

/**
 * Storage consumption percentage (0-100+) for a user, based on
 * storageUsedGB / storageQuotaGB. Returns 0 when no quota is set,
 * rather than dividing by zero.
 */
export function getStoragePercent(storageUsedGB: number, storageQuotaGB: number): number {
  if (!storageQuotaGB || storageQuotaGB <= 0) {
    return 0;
  }
  return (storageUsedGB / storageQuotaGB) * 100;
}

/**
 * Storage health thresholds, applied to storage consumption percentage:
 *   < 80%        -> Healthy
 *   80% - 89.99%  -> Warning
 *   90% - 94.99%  -> High Risk
 *   >= 95%        -> Critical
 */
export function getHealthLevel(storageUsedGB: number, storageQuotaGB: number): HealthLevel {
  const percent = getStoragePercent(storageUsedGB, storageQuotaGB);
  if (percent >= 95) {
    return 'Critical';
  }
  if (percent >= 90) {
    return 'High Risk';
  }
  if (percent >= 80) {
    return 'Warning';
  }
  return 'Healthy';
}
