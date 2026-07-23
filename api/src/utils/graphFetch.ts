import fetch, { RequestInit, Response } from 'node-fetch';

/**
 * Shared Microsoft Graph fetch helper with exponential backoff + jitter
 * and Retry-After support for HTTP 429 (Too Many Requests) responses.
 *
 * Used by every Graph call in this backend (reports + per-user profile
 * lookups) so that a large tenant (thousands of OneDrive owners) never
 * has to be enriched in bulk to avoid throttling - callers should also
 * minimize call volume (e.g. only calling per-user endpoints for
 * currently visible rows), but this helper protects whatever calls are
 * made from failing outright on a transient 429.
 *
 * Callers are also expected to de-duplicate concurrent requests for the
 * SAME resource (see graphReportsService.ts's pendingReportFetches / the
 * backend dashboard cache's in-flight promise) so that a 429 never
 * triggers multiple independent retry loops in parallel - only ONE
 * retry loop should ever be running per resource at a time.
 */

const MAX_RETRIES = 5;
const BASE_DELAY_MS = 1000;
const MAX_DELAY_MS = 30 * 1000;

/**
 * Thrown when Microsoft Graph keeps returning 429 after all retries have
 * been exhausted. Callers (Azure Function entry points) should catch this
 * specifically and return a friendly, non-retrying error response to the
 * client instead of a generic 502.
 */
export class GraphThrottledError extends Error {
  public readonly isThrottled = true;

  constructor(logContext: string) {
    super(`Microsoft Graph is throttling requests for ${logContext} after ${MAX_RETRIES} retries.`);
    this.name = 'GraphThrottledError';
  }
}

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

// Retry-After is honored, but capped so a single Graph-specified wait
// can never by itself exhaust the Azure Function's execution timeout
// (especially now that the dashboard downloads two reports back-to-back
// in a single request - see getOneDriveDashboardData.ts).
const MAX_RETRY_AFTER_SECONDS = 60;

function parseRetryAfterSeconds(response: Response): number | undefined {
  const header = response.headers.get('retry-after');
  if (!header) {
    return undefined;
  }
  const seconds = Number(header);
  return Number.isFinite(seconds) && seconds >= 0 ? Math.min(seconds, MAX_RETRY_AFTER_SECONDS) : undefined;
}

/**
 * Performs a fetch, automatically retrying on HTTP 429 using the
 * Retry-After response header when present, or exponential backoff
 * (with jitter) otherwise. Non-429 responses (including other error
 * statuses) are returned as-is for the caller to handle.
 *
 * Throws GraphThrottledError if still receiving 429 after MAX_RETRIES
 * attempts, rather than returning the (useless) final 429 response.
 */
export async function fetchWithRetry(
  url: string,
  init: RequestInit,
  logContext: string
): Promise<Response> {
  let attempt = 0;

  while (true) {
    const response = await fetch(url, init);

    if (response.status !== 429) {
      return response;
    }

    attempt++;
    if (attempt > MAX_RETRIES) {
      // eslint-disable-next-line no-console
      console.error(
        `[Graph throttling] ${logContext}: still receiving 429 after ${MAX_RETRIES} retries. Giving up.`
      );
      throw new GraphThrottledError(logContext);
    }

    const retryAfterSeconds = parseRetryAfterSeconds(response);
    const backoffMs = Math.min(MAX_DELAY_MS, BASE_DELAY_MS * Math.pow(2, attempt - 1));
    const jitterMs = Math.floor(Math.random() * 250);
    const delayMs = retryAfterSeconds !== undefined ? retryAfterSeconds * 1000 : backoffMs + jitterMs;
    const delaySeconds = Math.round(delayMs / 1000);

    // eslint-disable-next-line no-console
    console.warn(
      `Graph throttling detected: ${logContext}. Retrying after ${delaySeconds} seconds. ` +
      `[attempt ${attempt}/${MAX_RETRIES}]${retryAfterSeconds !== undefined ? ' (Retry-After header honored)' : ' (exponential backoff)'}`
    );

    await sleep(delayMs);
  }
}
