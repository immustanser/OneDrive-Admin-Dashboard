import { IOneDriveUser, OneDriveStatus, IStorageTrendPoint } from '../models';

const FIRST_NAMES: string[] = ['Olivia', 'Liam', 'Emma', 'Noah', 'Ava', 'Ethan', 'Sophia', 'Mason', 'Isabella', 'Lucas', 'Mia', 'Aiden', 'Amelia', 'Jackson', 'Harper', 'Elijah', 'Evelyn', 'James', 'Abigail', 'Benjamin'];
const LAST_NAMES: string[] = ['Johnson', 'Williams', 'Brown', 'Jones', 'Garcia', 'Miller', 'Davis', 'Rodriguez', 'Martinez', 'Wilson', 'Anderson', 'Taylor', 'Thomas', 'Moore', 'Jackson', 'Martin', 'Lee', 'Perez', 'Thompson', 'White'];
const DEPARTMENTS: string[] = ['Finance', 'Human Resources', 'Sales', 'Marketing', 'Engineering', 'IT', 'Legal', 'Operations', 'Customer Support', 'Executive'];

function seededRandom(seed: number): () => number {
  let s = seed;
  return () => {
    s = (s * 9301 + 49297) % 233280;
    return s / 233280;
  };
}

function pick<T>(arr: T[], rnd: () => number): T {
  return arr[Math.floor(rnd() * arr.length)];
}

function computeStatus(daysSinceActive: number, storagePct: number): OneDriveStatus {
  if (daysSinceActive > 90) {
    return 'Inactive';
  }
  if (daysSinceActive > 30 || storagePct > 90) {
    return 'Warning';
  }
  return 'Active';
}

export function generateMockOneDriveUsers(count: number = 250): IOneDriveUser[] {
  const rnd = seededRandom(42);
  const users: IOneDriveUser[] = [];

  for (let i = 0; i < count; i++) {
    const first = pick(FIRST_NAMES, rnd);
    const last = pick(LAST_NAMES, rnd);
    const displayName = `${first} ${last}`;
    const email = `${first}.${last}${i}@contoso.com`.toLowerCase();
    const department = pick(DEPARTMENTS, rnd);
    const quota = pick([1024, 2048, 5120], rnd);
    const daysSinceActive = Math.floor(rnd() * 140);
    const used = Math.round(rnd() * quota * 0.95 * 100) / 100;
    const storagePct = (used / quota) * 100;
    const status = computeStatus(daysSinceActive, storagePct);
    const lastActivity = new Date();
    lastActivity.setDate(lastActivity.getDate() - daysSinceActive);

    const managerIndex = i > 0 ? Math.floor(rnd() * i) : -1;
    const manager = managerIndex >= 0 ? users[managerIndex].displayName : (rnd() > 0.92 ? '' : 'Sarah Connor');

    users.push({
      id: `user-${i}`,
      displayName,
      email,
      department,
      jobTitle: pick(['Manager', 'Analyst', 'Director', 'Specialist', 'Coordinator', 'Associate'], rnd),
      oneDriveUrl: `https://contoso-my.sharepoint.com/personal/${first.toLowerCase()}_${last.toLowerCase()}_contoso_com`,
      storageUsedGB: used,
      storageQuotaGB: quota,
      filesCount: Math.floor(rnd() * 15000) + 50,
      lastActivityDate: lastActivity.toISOString(),
      manager,
      status,
      sharedFilesCount: Math.floor(rnd() * 300),
      externalSharedFilesCount: Math.floor(rnd() * 40),
      anonymousLinksCount: Math.floor(rnd() * 15),
      companyLinksCount: Math.floor(rnd() * 60),
      isActive30Days: daysSinceActive <= 30,
      isActive60Days: daysSinceActive <= 60,
      isActive90Days: daysSinceActive <= 90
    });
  }

  return users;
}

export function generateMockStorageTrend(months: number = 12): IStorageTrendPoint[] {
  const rnd = seededRandom(7);
  const points: IStorageTrendPoint[] = [];
  let base = 2000;
  const now = new Date();

  for (let i = months - 1; i >= 0; i--) {
    const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    base += rnd() * 150;
    points.push({
      month: d.toLocaleString('default', { month: 'short', year: '2-digit' }),
      storageUsedGB: Math.round(base * 100) / 100
    });
  }

  return points;
}
