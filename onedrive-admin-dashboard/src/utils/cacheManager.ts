interface ICacheEntry<T> {
  value: T;
  expiresAt: number;
}

export class CacheManager {
  private static store: Map<string, ICacheEntry<unknown>> = new Map();

  public static get<T>(key: string): T | undefined {
    const entry = this.store.get(key);
    if (!entry) {
      return undefined;
    }
    if (Date.now() > entry.expiresAt) {
      this.store.delete(key);
      return undefined;
    }
    return entry.value as T;
  }

  public static set<T>(key: string, value: T, ttlMs: number = 5 * 60 * 1000): void {
    this.store.set(key, { value, expiresAt: Date.now() + ttlMs });
  }

  public static clear(key?: string): void {
    if (key) {
      this.store.delete(key);
    } else {
      this.store.clear();
    }
  }

  public static async getOrFetch<T>(key: string, fetcher: () => Promise<T>, ttlMs: number = 5 * 60 * 1000): Promise<T> {
    const cached = this.get<T>(key);
    if (cached !== undefined) {
      return cached;
    }
    const value = await fetcher();
    this.set(key, value, ttlMs);
    return value;
  }
}
