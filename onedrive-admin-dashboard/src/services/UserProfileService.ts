import { GraphService } from './GraphService';

export interface IUserProfileInfo {
  department: string;
  jobTitle: string;
  manager: string;
}

const EMPTY_PROFILE: IUserProfileInfo = { department: '', jobTitle: '', manager: '' };

/**
 * In-memory cache of per-user profile lookups (department, jobTitle,
 * manager), keyed by userPrincipalName/email (case-insensitive).
 *
 * IMPORTANT: profiles are only ever fetched for users that are actually
 * visible on screen right now (see hooks/usePagedOneDriveUsers.ts), NEVER
 * for the full tenant. This is what avoids Microsoft Graph 429 (Too Many
 * Requests) on large tenants - e.g. a 25-row inventory page results in at
 * most 25 GET /api/user-profile calls, and revisiting a page already
 * fetched costs zero additional calls because of this cache.
 */
const cache = new Map<string, IUserProfileInfo>();
const pending = new Map<string, Promise<IUserProfileInfo>>();

// Tracks UPNs whose most recent profile lookup FAILED (network/HTTP
// error), as opposed to succeeding with genuinely empty fields. Used by
// the Inventory grid to show "Unknown" only on real lookup failures,
// never on a legitimately empty department/manager value.
const errorSet = new Set<string>();

export class UserProfileService {
  /**
   * Synchronous read of whatever profile info is already cached for this
   * user. Never triggers a network call - used by the Storage Analytics
   * charts, which should only use currently-loaded department values
   * rather than enriching the full tenant.
   */
  public static getCachedProfile(userPrincipalName: string): IUserProfileInfo | undefined {
    if (!userPrincipalName) {
      return undefined;
    }
    return cache.get(userPrincipalName.toLowerCase());
  }

  /**
   * Fetches (or returns the cached) profile for a single user.
   */
  public static async getUserProfile(userPrincipalName: string): Promise<IUserProfileInfo> {
    if (!userPrincipalName) {
      return EMPTY_PROFILE;
    }

    const key = userPrincipalName.toLowerCase();
    const cached = cache.get(key);
    if (cached) {
      return cached;
    }

    let inFlight = pending.get(key);
    if (!inFlight) {
      inFlight = GraphService.getUserProfile(userPrincipalName)
        .then((profile) => {
          errorSet.delete(key);
          cache.set(key, profile);
          pending.delete(key);
          return profile;
        })
        .catch(() => {
          // Best-effort enrichment: on failure, cache an empty profile so
          // we don't hammer the API again for this user on every render,
          // but don't block/break the inventory grid. Mark this UPN as a
          // lookup failure so the UI can show "Unknown" instead of a
          // blank cell (which would be indistinguishable from a
          // genuinely empty successful result).
          errorSet.add(key);
          cache.set(key, EMPTY_PROFILE);
          pending.delete(key);
          return EMPTY_PROFILE;
        });
      pending.set(key, inFlight);
    }
    return inFlight;
  }

  /**
   * True when the most recent lookup for this user FAILED (as opposed to
   * succeeding with empty fields). Used to render "Unknown" rather than a
   * blank cell.
   */
  public static isProfileError(userPrincipalName: string): boolean {
    if (!userPrincipalName) {
      return false;
    }
    return errorSet.has(userPrincipalName.toLowerCase());
  }

  /**
   * Fetches profiles only for the given (currently visible) user
   * principal names, skipping any already cached/in-flight. For a page
   * size of 25, this results in at most 25 Graph-backed API calls -
   * never one per user in the whole tenant.
   */
  public static async getUserProfiles(userPrincipalNames: string[]): Promise<Map<string, IUserProfileInfo>> {
    const unique = Array.from(new Set(userPrincipalNames.filter(Boolean)));
    await Promise.all(unique.map((upn) => UserProfileService.getUserProfile(upn)));

    const result = new Map<string, IUserProfileInfo>();
    unique.forEach((upn) => {
      const profile = cache.get(upn.toLowerCase());
      if (profile) {
        result.set(upn, profile);
      }
    });
    return result;
  }
}
