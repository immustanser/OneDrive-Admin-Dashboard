import { ClientSecretCredential, AccessToken } from '@azure/identity';

/**
 * SECURITY NOTE
 * -------------
 * TENANT_ID, CLIENT_ID, CLIENT_SECRET and GRAPH_SCOPE are read from
 * environment variables (`process.env`), which are populated from:
 *   - Local dev: api/local.settings.json  (git-ignored, never committed)
 *   - Azure:     Function App > Configuration > Application settings,
 *                or a Key Vault reference, e.g.
 *                CLIENT_SECRET = @Microsoft.KeyVault(SecretUri=https://<vault>.vault.azure.net/secrets/<name>/)
 *
 * The client secret NEVER appears in source code, NEVER gets returned to
 * the SPFx client, and must NEVER be committed to source control.
 */

let cachedCredential: ClientSecretCredential | undefined;
let cachedToken: AccessToken | undefined;

function getRequiredEnv(name: string): string {
  const value = process.env[name];
  if (!value) {
    throw new Error(
      `Missing required application setting "${name}". Configure it in ` +
      `api/local.settings.json for local development, or in Azure Function ` +
      `App Settings / Key Vault for production.`
    );
  }
  return value;
}

function getCredential(): ClientSecretCredential {
  if (!cachedCredential) {
    const tenantId = getRequiredEnv('TENANT_ID');
    const clientId = getRequiredEnv('CLIENT_ID');
    const clientSecret = getRequiredEnv('CLIENT_SECRET');
    cachedCredential = new ClientSecretCredential(tenantId, clientId, clientSecret);
  }
  return cachedCredential;
}

/**
 * Acquires (and caches in-memory for the lifetime of the Function host)
 * an app-only Microsoft Graph access token via the OAuth2 client
 * credentials flow.
 */
export async function getGraphAccessToken(): Promise<string> {
  const scope = process.env.GRAPH_SCOPE || 'https://graph.microsoft.com/.default';

  const nowMs = Date.now();
  const safetyMarginMs = 60 * 1000;
  if (cachedToken && cachedToken.expiresOnTimestamp - safetyMarginMs > nowMs) {
    return cachedToken.token;
  }

  const credential = getCredential();
  const token = await credential.getToken(scope);
  if (!token) {
    throw new Error('Failed to acquire Microsoft Graph access token via client credentials flow.');
  }
  cachedToken = token;
  return token.token;
}
