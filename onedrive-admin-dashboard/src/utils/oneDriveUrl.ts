/**
 * Derives a best-effort OneDrive personal-site URL for a user when
 * Microsoft Graph's OneDrive usage report leaves "Site URL" blank
 * (this happens for some rows depending on tenant reporting behavior -
 * it is not a parsing bug on our side).
 *
 * Falls back to the standard, publicly documented OneDrive personal-site
 * URL convention: https://<tenant>-my.sharepoint.com/personal/<sanitized upn>
 * This is a best-effort convention-based guess, not a guaranteed-correct
 * link - callers should only use it when the real oneDriveUrl is empty.
 */
export function buildOneDriveFallbackUrl(email: string, tenantRootUrl: string): string {
  if (!email || !tenantRootUrl) {
    return '';
  }

  let host: string;
  try {
    host = new URL(tenantRootUrl).hostname;
  } catch {
    return '';
  }

  // Only handle the standard "<tenant>.sharepoint.com" root pattern; if the
  // root URL doesn't look like a standard tenant root (e.g. a vanity/custom
  // domain), we cannot reliably guess the "-my" host, so bail out to "N/A".
  const match = /^([^.]+)\.sharepoint\.com$/i.exec(host);
  if (!match) {
    return '';
  }

  const tenantName = match[1];
  const sanitizedUpn = email.trim().replace(/[^a-zA-Z0-9]/g, '_');
  if (!sanitizedUpn) {
    return '';
  }

  return `https://${tenantName}-my.sharepoint.com/personal/${sanitizedUpn}`;
}
