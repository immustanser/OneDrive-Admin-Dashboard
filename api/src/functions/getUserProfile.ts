import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getUserProfile as fetchUserProfile } from '../services/graphUserProfileService';

/**
 * GET /api/user-profile?upn={userPrincipalName}
 *
 * Returns { department, manager, jobTitle } for a SINGLE user only.
 *
 * This endpoint exists so the SPFx Inventory grid can enrich only the
 * rows currently visible on screen (e.g. 25 calls for a page size of 25)
 * instead of the bulk dashboard/account-details endpoints enriching every
 * OneDrive owner in the tenant, which causes Microsoft Graph to respond
 * with 429 Too Many Requests on large tenants. Results are cached
 * in-memory per userPrincipalName (see graphUserProfileService.ts) and
 * Graph 429 responses are retried with exponential backoff / Retry-After
 * (see utils/graphFetch.ts).
 */
export async function getUserProfile(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  const upn = request.query.get('upn');
  if (!upn) {
    return {
      status: 400,
      jsonBody: { error: 'Missing required query parameter "upn".' }
    };
  }

  try {
    const profile = await fetchUserProfile(upn);
    return {
      status: 200,
      jsonBody: {
        department: profile.department,
        manager: profile.manager,
        jobTitle: profile.jobTitle
      }
    };
  } catch (error) {
    context.error('getUserProfile failed:', error);
    return {
      status: 502,
      jsonBody: {
        error: 'Unable to retrieve user profile from Microsoft Graph.',
        message: error instanceof Error ? error.message : 'Unknown error'
      }
    };
  }
}

app.http('getUserProfile', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'user-profile',
  handler: getUserProfile
});
